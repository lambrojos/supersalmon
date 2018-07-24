const pump = require('pump')
const XlsxStreamReader = require('xlsx-stream-reader')
const { Writable } = require('stream')
const through = require('through')
const FileType = require('stream-file-type')
const objectChunker = require('object-chunker')

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

    function write (row) {
      if (row.every(isEmpty)) { return true }
      const rowObj = {}
      for (let i = 0; i < cols.length; i++) {
        rowObj[cols[i]] = row[i]
      }
      return this.queue(rowObj)
    }

    const toStream = (reader) => {
      const s = through(write)
      s.on('drain', () => reader.workSheetStream.resume())
      reader.on('row', function onRow (row, i) {
        try {
          if (!cols) {
            if (reader.sheetData.dimension) {
              const lines = reader.sheetData.dimension[0].attributes.ref.match(/\d+$/)
              onLineCount(parseInt(lines, 10) - 1)
            }
            cols = row.values.map(mapColumns)
          } else if (!s.write(row.values)) {
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
        write (row, encoding, cb) {
          return processor(row, i).then(() => {
            i += 1 * chunkSize
            if (i === limit) {
              readStream.destroy()
              inputStream.destroy()
              return resolve(i)
            }
            return cb(null)
          }).catch(cb)
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
