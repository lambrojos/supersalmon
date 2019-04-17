/* eslint-env node, mocha */
const { expect } = require('chai')
const XLSXProcessor = require('..')
const { join } = require('path')
const { createReadStream } = require('fs')
const { Writable } = require('stream')

describe.only('XSLT Processor', () => {
  it('processes an xslt file.', async () => {
    const processed = await XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: async (data, i) => {
        expect(i).to.be.a('number')
        expect(data['first name']).to.be.ok
      }
    })
    expect(processed).to.equal(19)
  })

  it('processes an xslt file without headers.', async () => {
    const processed = await XLSXProcessor({
      hasHeaders: false,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'prova.xlsx'))
    }).processor({
      onRow: async data => {
        expect(Object.keys(data)).to.deep.equal(Array.from(Array(10).keys()).map(k => k.toString()))
      }
    })
    expect(processed).to.equal(7)
  })

  it('processes an xslt file without headers while providing a column mapoer', async () => {
    const processed = await XLSXProcessor({
      hasHeaders: false,
      mapColumns: (_v, i) => i * 2,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'prova.xlsx'))
    }).processor({
      onRow: async data => {
        expect(Object.keys(data)).to.deep.equal(Array.from(Array(10).keys()).map(k => (k * 2).toString()))
      }
    })
    expect(processed).to.equal(7)
  })

  it('allows chunked processing', async () => {
    const processed = await XLSXProcessor({
      chunkSize: 2,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'prova.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: async (data, i) => {
        expect(data).to.have.lengthOf(2)
        expect(i % 2).to.equal(0)
      }
    })
    expect(processed).to.equal(6)
  })

  it('formats by default', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'dates.xlsx'))
    await XLSXProcessor({
      inputStream,
      mapColumns: i => i
    }).processor({
      onRow: async data => {
        expect(data['Data di nascita']).to.match(/\d{2}\/\d{2}\/\d{4}/)
        expect(data['cap']).to.match(/\d{5}/)
      }
    })
  })

  it('can return formats', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'dates.xlsx'))
    await XLSXProcessor({
      inputStream,
      mapColumns: i => i,
      returnFormats: true
    }).processor({
      onRow: async data => {
        expect(data.values['Data di nascita']).to.match(/\d{2}\/\d{2}\/\d{4}/)
        expect(data.formats['Data di nascita']).to.equal('DD/MM/YYYY')
        expect(data.values['cap']).to.match(/\d{5}/)
      }
    })
  })

  it('can be used with predefined formats', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'predefined_formats.xlsx'))
    await XLSXProcessor({
      inputStream,
      mapColumns: i => i
    }).processor({
      onRow: async data => {
        expect(data['Data di nascita']).to.match(/\d\/\d{1,2}\/\d{2}/)
        expect(data['cap']).to.match(/\d{5}/)
      }
    })
  })

  it('can return predefined formats', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'predefined_formats.xlsx'))
    await XLSXProcessor({
      inputStream,
      mapColumns: i => i,
      returnFormats: true
    }).processor({
      onRow: async data => {
        expect(data.values['Data di nascita']).to.match(/\d\/\d{1,2}\/\d{2}/)
        expect(data.formats['Data di nascita']).to.equal('m/d/yy')
        expect(data.formats['cap']).to.equal('General')
      }
    })
  })

  it('can disable formatting', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'dates.xlsx'))
    await XLSXProcessor({
      inputStream,
      mapColumns: i => i,
      formatting: false
    }).processor({
      onRow: async data => {
        expect(data['Data di nascita']).to.match(/^\d{5}$/)
      }
    })
  })

  it('allows stream limit interruption', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'test.xlsx'))
    const processed = await XLSXProcessor({
      inputStream,
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: () => {},
      limit: 10
    })
    expect(processed).to.equal(10)
  })

  it('processes largeish files', async () => {
    await XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'big.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: async () => {
      }
    })
  })

  it('processes this other xslt file.', async () => {
    try {
      await XLSXProcessor({
        inputStream: createReadStream(join(__dirname, 'fixtures', 'broken.xlsx')),
        mapColumns: colName => colName.toLowerCase().trim()
      }).processor({
        onRow: async () => {}
      })
    } catch (e) {
      expect(e.message).to.equal('Unknown column ID')
    }
  })

  it('allows the user to access the row objects stream', () => {
    const stream = XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).stream()
    stream.pipe(new Writable({
      objectMode: true,
      write (chunk, _enc, cb) {
        cb()
      }
    }))
  })

  it('allows the user to process the first n records', cb => {
    let nRows = 0
    const stream = XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).stream(2)
    stream.pipe(new Writable({
      objectMode: true,
      write (chunk, _enc, cb) {
        nRows++
        cb()
      }
    }))
    stream.on('end', () => {
      expect(nRows).to.equal(2)
      cb()
    })
  })

  it('does not report an error if the first line is empty and headers are not required', async () => {
    await XLSXProcessor({
      hasHeaders: false,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'missingHeaders.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: () => {}
    })
  })

  it('reports an error if the first line is empty and headers are required', async () => {
    try {
      await XLSXProcessor({
        hasHeaders: true,
        inputStream: createReadStream(join(__dirname, 'fixtures', 'missingHeaders.xlsx')),
        mapColumns: colName => colName.toLowerCase().trim()
      }).processor({
        onRow: () => {}
      })
      throw new Error('Expected to throw')
    } catch (e) {
      expect(e.message).to.equal('Header row is empty')
    }
  })

  it('reports an error if the first line is partially empty and headers are required', async () => {
    try {
      await XLSXProcessor({
        hasHeaders: true,
        inputStream: createReadStream(join(__dirname, 'fixtures', 'missingcols.xlsx')),
        mapColumns: colName => colName.toLowerCase().trim()
      }).processor({})
      throw new Error('Expected to throw')
    } catch (e) {
      expect(e.message).to.equal('Missing column name at index 1')
    }
  })

  it('parses dates on this file', async () => {
    await XLSXProcessor({
      hasHeaders: true,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'noformatcodes.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).processor({
      onRow: data => {
        expect(data['data di nascita']).to.match(/\d\/\d{1,2}\/\d{2}/)
      }
    })
  })

  it('reports an error if the first line is partially empty and headers are required, while being chunky', async () => {
    try {
      await XLSXProcessor({
        hasHeaders: true,
        chunkSize: 2,
        inputStream: createReadStream(join(__dirname, 'fixtures', 'missingcols.xlsx')),
        mapColumns: colName => colName.toLowerCase().trim()
      }).processor({
      })
      throw new Error('Expected to throw')
    } catch (e) {
      expect(e.message).to.equal('Missing column name at index 1')
    }
  })

  it('reports an error on a non xlsx-file', async () => {
    try {
      await XLSXProcessor({
        inputStream: createReadStream(join(__dirname, 'fixtures', 'notanxlsx.xlsx')),
        mapColumns: colName => colName.toLowerCase().trim()
      }).processor({
        onRow: async () => {}
      })
    } catch (e) {
      expect(e.message).to.equal('Invalid file type')
    }
  })

  it('does not eat records in case of very slow readable streams', cb => {
    let lastChunk = null
    const stream = XLSXProcessor({
      lenientColumnMapping: true,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'slow.xlsx')),
      mapColumns: colName => colName.toLowerCase().trim()
    }).stream()
    const slow = new Writable({
      objectMode: true,
      write (chunk, _enc, cb) {
        console.log(chunk)
        lastChunk = chunk
        setTimeout(() => {
          cb()
        }, 500)
      }
    })
    stream.pipe(slow)
    slow.on('finish', () => {
      expect(lastChunk['first name']).to.equal('Minoru')
      cb()
    })
  })
})
