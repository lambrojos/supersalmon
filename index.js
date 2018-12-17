const pump = require('pump')
const XlsxStreamReader = require('xlsx-stream-reader')
const { Writable, Transform } = require('stream')
const FileType = require('stream-file-type')
const objectChunker = require('object-chunker')
const debug = require('debug')('salmon')

const isEmpty = val => val === undefined || val === null || val === ''
const byIndex = (_val, i) => i

const transformRow = (row, cols) => {
  const rowObj = {}
  for (let i = 0; i < cols.length; i++) {
    rowObj[cols[i]] = row.values[i]
  }
  return rowObj
}

const withFormats = (row, cols) => {
  row.formats.shift()
  const formats = {}
  for (let i = 0; i < cols.length; i++) {
    formats[cols[i]] = row.formats[i]
  }
  return {
    values: transformRow(row, cols),
    formats
  }
}
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
  mapColumns = byIndex,
  onLineCount = () => {},
  limit = Infinity,
  returnFormats = false,
  formatting = true,
  chunkSize = 1,
  hasHeaders = true
}) => {
  let cols = null
  let detected = false
  let stream
  let detector
  let reader
  const onErr = err => {
    if (!err) return
    debug(err)
    inputStream.destroy()
    detector.destroy()
    if (reader) { reader.removeAllListeners('row') }
    if (stream) { stream.destroy(err) }
  }

  const rowTransformer = returnFormats ? withFormats : transformRow

  return {
    stream () {
      detector = new FileType()

      stream = new Transform({
        objectMode: true,
        transform (chunk, enc, cb) {
          if (chunk.values.every(isEmpty)) { return cb() }
          this.push(rowTransformer(chunk, cols))
          cb()
        }
      })

      const checkAndPipe = (fileType) => {
        if (!fileType || fileType.mime !== 'application/zip') {
          onErr(new Error('Invalid file type'))
        } else {
          pump(detector, workBookReader, onErr)
          workBookReader.on('end', () => {
            if (!detected) {
              onErr(new Error('Invalid file type'))
            }
          })
        }
      }

      const readSheet = workSheetReader => {
        detected = true
        reader = workSheetReader
        workSheetReader.workSheetStream.on('error', onErr)
        // read only the first worksheet for now
        if (workSheetReader.id > 1) { workSheetReader.skip(); return }
        // worksheet reader is an event emitter - we have to convert it to a read stream
        // signal stream end when the event emitter is finished
        workSheetReader.on('end', () => stream.push(null))
        workSheetReader.process()

        stream.on('drain', () => {
          debug('resume stream')
          workSheetReader.workSheetStream.resume()
        })
        workSheetReader.on('row', (row, i) => {
          row.values.shift()
          debug('row received')
          try {
            if (row.values.every(isEmpty)) {
              if (!cols && hasHeaders) { throw new Error('Header row is empty') } else { return }
            }
            if (!cols) {
              if (workSheetReader.sheetData.dimension) {
                const lines = workSheetReader.sheetData.dimension[0].attributes.ref.match(/\d+$/)
                onLineCount(parseInt(lines, 10) - 1)
              }
              cols = row.values.map(mapColumns)
              if (!hasHeaders) {
                stream.write(row)
              } else if (row.values.every(isEmpty)) {
                throw new Error('Empty header row')
              }
            } else if (!stream.write(row)) {
              debug('pausing stream')
              workSheetReader.workSheetStream.pause()
            }
          } catch (err) {
            onErr(err)
          }
        })
      }
      const workBookReader = new XlsxStreamReader({ formatting, returnFormats })
      workBookReader.on('error', onErr)
      workBookReader.on('worksheet', readSheet)

      detector.on('file-type', checkAndPipe)
      pump(inputStream, detector, onErr)

      return chunkSize > 1 ? pump(stream, objectChunker(chunkSize)) : stream
    },
    processor ({ onRow, limit }) {
      let i = 0
      const readStream = this.stream()
      return new Promise((resolve, reject) => {
        const processRow = new Writable({
          objectMode: true,
          async write (row, encoding, cb) {
            // if this fails, pump will cleanup the broken streams.
            try {
              await onRow(row, i)
              i += chunkSize
              if (i === limit) {
                readStream.destroy()
                detector.destroy()
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
        pump(readStream, processRow, err => {
          onErr(err)
          if (err) {
            reject(err)
          } else {
            resolve(i)
          }
        })
      })
    }
  }
}
