const { expect } = require('chai');
const XLSXProcessor = require('..');
const { join } = require('path')
const { createReadStream } = require('fs')

describe('XSLT Processor', () => {
  it.only('processes an xslt file.', async () => {
    const processed = await XLSXProcessor(
      createReadStream(join(__dirname, 'fixtures', 'error.xlsx')),
      (data) => {
        expect(data['first name']).to.be.ok;
      },
      colName => colName.toLowerCase().trim(),
      () => {},
    );
    expect(processed).to.equal(19);
  });

  it.only('processes this other xslt file.', async () => {
    try {
      await XLSXProcessor(
        createReadStream(join(__dirname, 'fixtures', 'broken.xlsx')),
        () => {},
        colName => colName.toLowerCase().trim(),
        () => {},
      );
    } catch (e) {
      expect(e.message).to.equal('Unknown column ID');
    }
  });
});
