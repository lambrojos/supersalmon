# Supersalmon

## Features
* Process xlsx files while streaming
* Records are processed one at a time by a user supplied promise returning function
* Strives for memory efficiency - Automatic backpressure handling
* Line count support (not all xlsx files contain this metadata)
* Can transform column names
* Will reject on non xlsx data
* Battle tested in production code
* Skips empty rows (maybe it's a caveat)
* Limit parsing: process a file to a specified row, then stop
* Allows to process rows in chunks (for multiple inserts)

## Caveats
* Only parses the first on sheet in a workbook
* Requires the first row to contain column names
* Skips empty rows (maybe it's a feature)

## Example Usage

```javascript
  const supersalmon = require('supersalmon')

  // The promise is resolved when the stream is completely processed
  // It resolves to the count of processed rows
  const processed = await supersalmon({

    // (required) Any readable stream will work - remember that only the first sheet will be parsed
    inputStream: createReadStream('huge.xlsx'),

    hasHeaders: true,

    // optional - returns data alongside their formats ( such as 'DD/MM/YYYY' for dates) - see tests
    hasFormats: false,

    // (required) transform column names- column names will become the key names of the processed objects
    mapColumns: cols => colName => colName.toLowerCase().trim(),

    // (optional) the last function is called when the line count metadata is encountered in the stream
    onLineCount: lineCount => notifyLineCount(lineCount),

    // enable or disable the underlying xlsx-stream-reader formatting feature
    formatting: false,

    // if true, does not throw an error when a cell does not have corresponding column name and skips the cell instead
    lenientColumnMapping: false

    // Returns rows in arrays of 3 elements
    chunkSize: 3
  }).processor({
    // (optional) parse until the 10th line then destroy streams and return
    limit: 10,
    // the row index tis provided as the second parameter
    onRow: ({name, surname}, i) => {
      doSomethingWithRowIndex(i)
      repository.insert({ name, surname })
    },
  });

  // Alternatively it is possibile to access directly the object stream

  const stream = supersalmon({ /* config */ }).stream()

  // it is possibile to take only the first n records
  const stream10 = supersalmon({ /* config */ }).stream(10)


  stream.pipe(myOtherStream)
```

Or see tests

## TODO List
- [ ] Write some real documentation
- [ ] Refactor
- [x] Add linter
- [ ] Add more tests
- [x] Support files without column names in the first row
- [ ] Parallel record processing
- [ ] Port to typescript
- [x] Better API
- [ ] Parse all sheets in a workbook

Issues and PRs more than welcome!


## Credits
A large part of the code base has been adapted from [xlsx-stream-reader](https://www.npmjs.com/package/xlsx-stream-reader)