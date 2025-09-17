// Copy-pasted from:
// https://github.com/catamphetamine/serverless-functions/blob/master/source/deploy/archive.js

// Uses `archiver` library.
// https://www.npmjs.com/package/archiver

// https://www.archiverjs.com/docs/archiver
import archiver from 'archiver'
// import { WritableStream } from 'memory-streams'
import fs from 'fs'

// `PassThrough` is available since Node version `0.10`.
import { PassThrough } from 'stream'

/**
 * A server-side *.zip archive creator.
 */
export default class Archive {
  constructor(outputPath) {
    if (outputPath) {
      this.outputStream = fs.createWriteStream(outputPath)
    } else {
      // // Won't work for memory streams.
      // // https://github.com/archiverjs/node-archiver/issues/336
      // this.outputStream = new WritableStream()
      this.outputStream = new PassThrough()
    }

    const archive = archiver('zip', {
      // // Sets the compression level.
      // zlib: { level: 9 }
    })

    this.archive = archive

    // `archive` has a `.pipe()` method which kinda makes it usable as a readable stream.
    // Although it's still not a "proper" implementation of a `Readable` stream.
    // https://github.com/archiverjs/node-archiver/issues/765
    // To get a "proper" implementation of a `Readable` stream, people use a workaround:
    // `const passThroughStream = new PassThrough()` and then `archive.pipe(passThroughStream)`.

    // An alternative way of how to create a `ReadableStream` from an `archive`.
    // https://github.com/archiverjs/node-archiver/issues/759
    //
    // const readableStream = new ReadableStream({
    //   start(controller) {
    //     archive.on('data', chunk => controller.enqueue(chunk));
    //     archive.on('end', () => controller.close());
    //     archive.on('error', error => controller.error(error));
    //
    //     archive.append(...);
    //     archive.append(...);
    //     archive.finalize();
    //   }
    // });

    this.promise = new Promise((resolve, reject) => {
      // listen for all archive data to be written
      // 'close' event is fired only when a file descriptor is involved
      this.outputStream.on('close', () => {
        resolve({ size: archive.pointer() })
      })

      // // This event is fired when the data source is drained no matter what was the data source.
      // // It is not part of this library but rather from the NodeJS Stream API.
      // // @see: https://nodejs.org/api/stream.html#stream_event_end
      // archive.on('end', function() {
      //   console.log('Data has been drained')
      //   resolve({
      //     // output: outputPath ? undefined : this.outputStream.toBuffer(),
      //     size: archive.pointer()
      //   })
      // })

      // good practice to catch warnings (ie stat failures and other non-blocking errors)
      archive.on('warning', function(error) {
        if (error.code === 'ENOENT') {
          // log warning
          console.warn(error)
        } else {
          reject(error)
        }
      })

      // good practice to catch this error explicitly
      archive.on('error', reject)

      // pipe archive data to the file
      archive.pipe(this.outputStream)
    })
  }

  file(filePath, internalPath) {
    this.archive.file(filePath, { name: internalPath })
  }

  directory(directoryPath, internalPath) {
    this.archive.directory(directoryPath, internalPath);
  }

  append(content, internalPath) {
    this.archive.append(content, { name: internalPath })
  }

  write() {
    // `.finalize()` returns some kind of `Promise` but it's not meant to be `await`ed.
    // https://github.com/archiverjs/node-archiver/issues/772
    this.archive.finalize()
  }

  getStream() {
    return this.outputStream
  }

  getPromise() {
    return this.promise
  }
}
