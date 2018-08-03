const { expect } = require('chai')
const XLSXProcessor = require('..')
const { join } = require('path')
const Promise = require('bluebird')
const { createReadStream } = require('fs')

describe('XSLT Processor', () => {
  it('processes an xslt file.', async () => {
    const processed = await XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      processor: async (data, i) => {
        expect(i).to.be.a('number')
        // console.log('DAAAAATAA', data)
        expect(data['first name']).to.be.ok
      },
      mapColumns: colName => colName.toLowerCase().trim()
    })
    expect(processed).to.equal(19)
  })

  it('allows chunked processing', async () => {
    const processed = await XLSXProcessor({
      chunkSize: 2,
      inputStream: createReadStream(join(__dirname, 'fixtures', 'prova.xlsx')),
      processor: async (data, i) => {
        expect(data).to.have.lengthOf(2)
        expect(i % 2).to.equal(0)
      },
      mapColumns: colName => colName.toLowerCase().trim()
    })
    expect(processed).to.equal(6)
  })

  it('formats by default', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'dates.xlsx'))
    const processed = await XLSXProcessor({
      inputStream,
      processor: async (data, i) => {
        expect(data['Data di nascita']).to.match(/\d{2}\/\d{2}\/\d{4}/)
        expect(data['cap']).to.match(/\d{5}/)
      },
      mapColumns: i => i
    })
  })

  it('can be used with predefined formats', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'predefined_formats.xlsx'))
    const processed = await XLSXProcessor({
      inputStream,
      processor: async (data, i) => {
        expect(data['Data di nascita']).to.match(/\d\/\d{1,2}\/\d{2}/)
        expect(data['cap']).to.match(/\d{5}/)
      },
      mapColumns: i => i
    })
  })

  it('can disable formatting', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'dates.xlsx'))
    const processed = await XLSXProcessor({
      inputStream,
      processor: async (data, i) => {
        expect(data['Data di nascita']).to.match(/^\d{5}$/)
      },
      mapColumns: i => i,
      formatting: false
    })
  })

  it('allows stream limit interruption', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'test.xlsx'))
    const processed = await XLSXProcessor({
      inputStream,
      processor: async (data, i) => {
      },
      mapColumns: colName => colName.toLowerCase().trim(),
      limit: 10
    })
    expect(processed).to.equal(10)
  })

  it('processes largeish files', async () => {
    await XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'big.xlsx')),
      processor: async () => {
      },
      mapColumns: colName => colName.toLowerCase().trim()
    })
  })

  it('processes this other xslt file.', async () => {
    try {
      await XLSXProcessor({
        inputStream: createReadStream(join(__dirname, 'fixtures', 'broken.xlsx')),
        processor: async () => {},
        mapColumns: colName => colName.toLowerCase().trim()
      })
    } catch (e) {
      expect(e.message).to.equal('Unknown column ID')
    }
  })
})
