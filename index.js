
/* const Excel = require('exceljs');
const Promise = require('bluebird'); */
const pump = require('pump');
const XlsxStreamReader = require('xlsx-stream-reader');
const { Writable } = require('stream');
const through = require('through');
/* const Promise = require('bluebird'); */
const FileType = require('stream-file-type');

const isEmpty = val => val === undefined || val === null || val === '';

/**
 * Processes a stream containing an XLSX file.
 * Calls the provided async `processor` function.
 * The processor function must handle its own error, unhandled exceptions will
 * cause the processing operation to fail.
 * Returns a promise containing the number of processed rows
 */
module.exports = (inputStream, processor, mapColumns, onLineCount) =>
  new Promise((resolve, reject) => {
    let i = 0;
    let cols = null;
    let detected = null;

    function write(row) {
      if (row.every(isEmpty)) { return true; }
      const rowObject = cols.reduce((memo, colValue, colIndex) =>
        Object.assign(memo, { [colValue]: row[colIndex] }),
      {});
      return this.queue(rowObject);
    }

    const workBookReader = new XlsxStreamReader();
    workBookReader.on('error', err => reject(err));

    workBookReader.on('worksheet', (workSheetReader) => {
      detected = true;

      workSheetReader.workSheetStream.on('error', (err) => {
        reject(err);
      });

      function toStream(ev) {
        const s = through(write);
        s.on('drain', () => workSheetReader.workSheetStream.resume());
        ev.on('row', function onRow(row) {
          try {
            if (!cols) {
              if (workSheetReader.sheetData.dimension) {
                const lines = workSheetReader.sheetData.dimension[0].attributes.ref.match(/\d+$/);
                onLineCount(parseInt(lines, 10) - 1);
              }
              cols = row.values.map(mapColumns)
            } else if (!s.write(row.values)) {
              ev.workSheetStream.pause();
            }
          } catch (err) {
            inputStream.destroy(err);
            reject(err);
            ev.removeListener('row', onRow);
          }
        });
        return s;
      }

      // read only the first worksheet for now
      if (workSheetReader.id > 1) { workSheetReader.skip(); return; }
      const processRow = new Writable({
        objectMode: true,
        async write(row, encoding, cb) {
        // if this fails, pump will cleanup the broken streams.
          try {
            await processor(row, i);
          } catch (err) {
            return cb(err);
          }
          i += 1;
          return cb(null);
        },
      });
      // worksheet reader is an event emitter - we have to convert it to a read stream
      const readStream = toStream(workSheetReader);
      // signal stream end when the event emitter is finished
      workSheetReader.on('end', () => { readStream.push(null); });
      // chunk and process rows
      pump(readStream, processRow, (err) => {
        if (err) { return reject(err); }
        return resolve(i);
      });
      workSheetReader.process();
    });

    const detector = new FileType();

    detector.on('file-type', (fileType) => {
      if (!fileType || fileType.mime !== 'application/zip') {
        inputStream.destroy();
        detector.destroy();
        reject(new Error('Invalid file type'));
      } else {
        pump(detector, workBookReader, (err) => {
          if (err) { reject(err); }
        });
        workBookReader.on('end', () => {
        // if we finished and did not detect any worksheet, then this zip
        // is not a xlsx
          if (!detected) {
            reject(new Error('Invalid file type'));
          }
        });
      }
    });

    pump(inputStream, detector, (err) => {
      if (err) { reject(err); }
    });
  });


/*
  // In case streaming does not work
  const workbook = new Excel.Workbook();
  const stream = workbook.xlsx.createInputStream();
  stream.on('done', () => {
    // console.log('OK', workbook);
    const used = process.memoryUsage().heapUsed / 1024 / 1024;
    console.log(`The script uses approximately ${used} MB`);
  });
  pump(inputStream, stream); */
