const { expect } = require('chai');
const XLSXProcessor = require('..');
const { join } = require('path')
const Promise = require('bluebird')
const { createReadStream } = require('fs')

describe('XSLT Processor', () => {
  it('processes an xslt file.', async () => {
    const processed = await XLSXProcessor({
      inputStream: createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      processor: (data, i) => {
        expect(i).to.be.a('number')
        expect(data['first name']).to.be.ok;
      },
      mapColumns: colName => colName.toLowerCase().trim(),
    });
    expect(processed).to.equal(19);
  });


  it('allows stream limit interruption', async () => {
    const inputStream = createReadStream(join(__dirname, 'fixtures', 'test.xlsx'));
    const processed = await XLSXProcessor({
      inputStream,
      processor: async (data, i) => {
        await Promise.delay(50)
      },
      mapColumns: colName => colName.toLowerCase().trim(),
      limit: 10
    });
    expect(processed).to.equal(10)
  });

  it('processes this other xslt file.', async () => {
    try {
      await XLSXProcessor({
        inputStream: createReadStream(join(__dirname, 'fixtures', 'broken.xlsx')),
        processor: () => {},
        mapColumns: colName => colName.toLowerCase().trim(),
      });
    } catch (e) {
      expect(e.message).to.equal('Unknown column ID');
    }
  });
});
