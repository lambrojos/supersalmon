const pump = require('pump')
const XlsxStreamReader = require('xlsx-stream-reader')
const { Writable, Transform } = require('stream')
const FileType = require('stream-file-type')
const objectChunker = require('object-chunker')
const debug = require('debug')('salmon')

const isEmpty = val => val === undefined || val === null || val === ''

/**
 * Processes a stream containing an XLSX file.
 * Calls the provided async `processor` function.
 * The processor function must handle its own error, unhandled exceptions will
 * cause the processing operation to fail.
 * Returns a promise containing the number of processed rows
 */
module.exports = ({
  inputStream,
  processor,
  mapColumns,
  onLineCount = () => {},
  limit = Infinity,
  formatting = true,
  chunkSize = 1
}) =>
  new Promise((resolve, reject) => {
    let i = 0
    let cols = null
    let detected = null

    const toStream = reader => {
      const s = new Transform({
        objectMode: true,
        transform (chunk, enc, cb) {
          if (chunk.every(isEmpty)) { return cb() }
          const rowObj = {}
          for (let i = 0; i < cols.length; i++) {
            rowObj[cols[i]] = chunk[i]
          }
          this.push(rowObj)
          cb()
        }
      })
      s.on('drain', () => {
        debug('resume stream')
        reader.workSheetStream.resume()
      })
      reader.on('row', function onRow (row, i) {
        debug('row received')
        try {
          if (!cols) {
            if (reader.sheetData.dimension) {
              const lines = reader.sheetData.dimension[0].attributes.ref.match(/\d+$/)
              onLineCount(parseInt(lines, 10) - 1)
            }
            cols = row.values.map(mapColumns)
          } else if (!s.write(row.values)) {
            debug('pausing stream')
            reader.workSheetStream.pause()
          }
        } catch (err) {
          inputStream.destroy(err)
          reject(err)
          reader.removeListener('row', onRow)
        }
      })
      return s
    }

    const readSheet = (workSheetReader) => {
      detected = true
      workSheetReader.workSheetStream.on('error', reject)
      // read only the first worksheet for now
      if (workSheetReader.id > 1) { workSheetReader.skip(); return }
      const readStream = toStream(workSheetReader)
      const processRow = new Writable({
        objectMode: true,
        async write (row, encoding, cb) {
        // if this fails, pump will cleanup the broken streams.
          try {
            await processor(row, i)
            i += chunkSize
            if (i === limit) {
              readStream.destroy()
              inputStream.destroy()
              this.destroy()
              // a premature close error will be raised.. this could actually be good as pump
              // will unpipe and cleanup
              return resolve(i)
            }
          } catch (err) {
            return cb(err)
          }
          return cb(null)
        }
      })
      // worksheet reader is an event emitter - we have to convert it to a read stream
      // signal stream end when the event emitter is finished
      workSheetReader.on('end', () => readStream.push(null))
      workSheetReader.process()
      const pipeline = chunkSize > 1
        ? [readStream, objectChunker(chunkSize), processRow]
        : [readStream, processRow]

      pump(pipeline, (err) => err ? reject(err) : resolve(i))
    }

    const checkAndPipe = (fileType) => {
      if (!fileType || fileType.mime !== 'application/zip') {
        inputStream.destroy()
        detector.destroy()
        reject(new Error('Invalid file type'))
      } else {
        pump(detector, workBookReader, (err) => {
          if (err) { reject(err) }
        })
        workBookReader.on('end', () => {
          if (!detected) {
            reject(new Error('Invalid file type'))
          }
        })
      }
    }

    const workBookReader = new XlsxStreamReader({ formatting })
    workBookReader.on('error', reject)
    workBookReader.on('worksheet', readSheet)

    const detector = new FileType()
    detector.on('file-type', checkAndPipe)
    pump(inputStream, detector, (err) => {
      if (err) { reject(err) }
    })
  })
