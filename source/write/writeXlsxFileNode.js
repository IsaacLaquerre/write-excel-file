import fs from 'fs'
import path from 'path'
import os from 'os'
import Stream, { Readable } from 'stream'

import Archive from './archive.js'
import getImageFileName from './getImageFileName.js'

import generateWorkbookXml from './files/workbook.xml.js'
import generateWorkbookXmlRels from './files/workbook.xml.rels.js'
import rels from './files/rels.js'
import generateContentTypesXml from './files/[Content_Types].xml.js'
import generateDrawingXml from './files/drawing.xml.js'
import generateDrawingXmlRels from './files/drawing.xml.rels.js'
import generateSheetXmlRels from './files/sheet.xml.rels.js'
import generateSharedStringsXml from './files/sharedStrings.xml.js'
import generateStylesXml from './files/styles.xml.js'

import { generateSheets } from './writeXlsxFile.common.js'

// This function doesn't use `async`/`await` in order to avoid adding `@babel/runtime` to `dependencies`.
// https://gitlab.com/catamphetamine/write-excel-file/-/issues/105
export default function writeXlsxFile(data, {
	filePath,
	buffer,
	sheet: sheetName,
	sheets: sheetNames,
	schema,
	columns,
	images,
	headerStyle,
	getHeaderStyle,
	fontFamily,
	fontSize,
	orientation,
	stickyRowsCount,
	stickyColumnsCount,
	showGridLines,
	zoomScale,
	conditionalStyles,
	rightToLeft,
	dateFormat
} = {}) {
	const {
		sheets,
		getSharedStrings,
		getStyles
	} = generateSheets({
		data,
		sheetName,
		sheetNames,
		schema,
		columns,
		images,
		headerStyle,
		getHeaderStyle,
		fontFamily,
		fontSize,
		orientation,
		stickyRowsCount,
		stickyColumnsCount,
		showGridLines,
		zoomScale,
		conditionalStyles,
		rightToLeft,
		dateFormat
	})

	return createDirectories().then(({
		rootPath,
		xlPath,
		mediaPath,
		drawingsPath,
		drawingsRelsPath,
		relsPath,
		worksheetsPath,
		worksheetsRelsPath
	}) => {
		const promises = [
			writeFile(path.join(relsPath, 'workbook.xml.rels'), generateWorkbookXmlRels({ sheets })),
			writeFile(path.join(xlPath, 'workbook.xml'), generateWorkbookXml({ sheets, stickyRowsCount, stickyColumnsCount })),
			writeFile(path.join(xlPath, 'styles.xml'), generateStylesXml(getStyles())),
			writeFile(path.join(xlPath, 'sharedStrings.xml'), generateSharedStringsXml(getSharedStrings()))
		]

		for (const { id, data, images } of sheets) {
			promises.push(writeFile(path.join(worksheetsPath, `sheet${id}.xml`), data))
			promises.push(writeFile(path.join(worksheetsRelsPath, `sheet${id}.xml.rels`), generateSheetXmlRels({ id, images })))
			if (images) {
				promises.push(writeFile(path.join(drawingsPath, `drawing${id}.xml`), generateDrawingXml({ images })))
				promises.push(writeFile(path.join(drawingsRelsPath, `drawing${id}.xml.rels`), generateDrawingXmlRels({ images, sheetId: id })))
				// Copy images to `xl/media` folder.
				for (const image of images) {
					const imageContentReadableStream = getReadableStream(image.content)
					const imageFilePath = path.join(mediaPath, getImageFileName(image, { sheetId: id, sheetImages: images }))
					promises.push(writeFileFromStream(imageFilePath, imageContentReadableStream))
				}
			}
		}

		return Promise.all(promises).then(() => {
			const archive = createArchive(filePath, {
				sheets,
				images,
				xlDirectoryPath: xlPath
			});

			if (filePath) {
				// Doesn't return anything.
				return archive.getPromise().then(() => {
					return removeDirectoryWithLegacyNodeVersionsSupport(rootPath)
				}).then(() => {
					// Doesn't return anything.
				})
			} else if (buffer) {
				// Returns a `Buffer`.
				return streamToBuffer(archive.getStream())
			} else {
				// Returns a readable `Stream`.
				return archive.getStream()
			}
		})
	})
}

// Creates a `*.zip` archive from Excel data and returns a readable `Stream`.
function createArchive(outputFilePath, {
	sheets,
	images,
	xlDirectoryPath
}) {
	// I dunno why `Archive` class is used here instead of something like `JSZip`.
	// `JSZip` is already used in `writeXlsxFileBrowser.js`, so maybe it would've also fit here.
	// In that case, `Archive` class could potentially be removed.
	const archive = new Archive(outputFilePath)

	archive.directory(xlDirectoryPath, 'xl')

	archive.append(rels, '_rels/.rels')

	archive.append(generateContentTypesXml({ sheets, images }), '[Content_Types].xml')

	archive.write()

	return archive
}

// According to Node.js docs:
// https://nodejs.org/api/fs.html#fswritefilefile-data-options-callback
// `contents` argument could be of type:
// * string — File path
// * Buffer
// * TypedArray
// * DataView
function writeFile(path, contents) {
	return new Promise((resolve, reject) => {
		fs.writeFile(path, contents, 'utf-8', (error) => {
			if (error) {
				return reject(error)
			}
			resolve()
		})
	})
}

function createDirectory(path) {
	return new Promise((resolve, reject) => {
		fs.mkdir(path, (error) => {
			if (error) {
				return reject(error)
			}
			resolve(path)
		})
	})
}

function createTempDirectory() {
	return new Promise((resolve, reject) => {
		fs.mkdtemp(path.join(os.tmpdir(), 'write-excel-file-'), (error, directoryPath) => {
			if (error) {
				return reject(error)
			}
			resolve(directoryPath)
		})
	})
}

function removeDirectoryWithLegacyNodeVersionsSupport(path) {
	if (fs.rm) {
		return removeDirectory(path)
	} else {
		removeDirectoryLegacySync(path)
  	return Promise.resolve()
	}
}

// `fs.rm()` is available in Node.js since `14.14.0`.
function removeDirectory(path) {
	return new Promise((resolve, reject) => {
		fs.rm(path, { recursive: true, force: true }, (error) => {
			if (error) {
				return reject(error)
			}
			resolve()
		})
	})
}

// For Node.js versions below `14.14.0`.
function removeDirectoryLegacySync(directoryPath) {
  const childNames = fs.readdirSync(directoryPath)
  for (const childName of childNames) {
    const childPath = path.join(directoryPath, childName)
    const stats = fs.statSync(childPath)
    if (childPath === '.' || childPath === '..') {
      // Skip.
    } else if (stats.isDirectory()) {
      // Remove subdirectory recursively.
      removeDirectoryLegacySync(childPath)
    } else {
      // Remove file.
      fs.unlinkSync(childPath)
    }
  }
  fs.rmdirSync(directoryPath)
}

// https://stackoverflow.com/a/67729663
function streamToBuffer(stream) {
	return new Promise((resolve, reject) => {
		const chunks = []
		stream.on('data', (chunk) => chunks.push(chunk))
		stream.on('end', () => resolve(Buffer.concat(chunks)))
		stream.on('error', reject)
	})
}

function copyFile(fromPath, toPath) {
	return new Promise((resolve, reject) => {
		fs.copyFile(fromPath, toPath, (error) => {
			if (error) {
				return reject(error)
			}
			resolve()
		})
	})
}

function getReadableStream(source) {
	if (source instanceof Stream) {
		return source
	}
	if (source instanceof Buffer) {
		return Readable.from(source)
	}
	if (typeof source === 'string') {
		return fs.createReadStream(source)
	}
	throw new Error('Unsupported content source: couldn\'t convert it to a readable stream')
}

function writeFileFromStream(filePath, readableStream) {
	const writableStream = fs.createWriteStream(filePath)
	readableStream.pipe(writableStream)
	return new Promise(resolve => writableStream.on('finish', resolve))
}

function createDirectories() {
	// There doesn't seem to be a way to just append a file into a subdirectory
	// in `archiver` library, hence using a hacky temporary directory workaround.
	// https://www.npmjs.com/package/archiver
	return createTempDirectory().then((rootPath) => {
		const xlPath = path.join(rootPath, 'xl')
		const mediaPath = path.join(xlPath, 'media')
		const drawingsPath = path.join(xlPath, 'drawings')
		const drawingsRelsPath = path.join(drawingsPath, '_rels')
		const relsPath = path.join(xlPath, '_rels')
		const worksheetsPath = path.join(xlPath, 'worksheets')
		const worksheetsRelsPath = path.join(worksheetsPath, '_rels')

		const directories = [
			xlPath,
			mediaPath,
			drawingsPath,
			drawingsRelsPath,
			relsPath,
			worksheetsPath,
			worksheetsRelsPath
		]

		const createAllDirectoriesPromise = directories.reduce((promise, directory) => {
			return promise.then(() => createDirectory(directory))
		}, Promise.resolve())

		return createAllDirectoriesPromise.then(() => {
			return {
				rootPath,
				xlPath,
				mediaPath,
				drawingsPath,
				drawingsRelsPath,
				relsPath,
				worksheetsPath,
				worksheetsRelsPath
			}
		})
	})
}