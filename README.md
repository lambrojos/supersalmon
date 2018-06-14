# XLSX stream processor

## Features
* Process xlsx files while streaming
* Records are processed one at a time by a user supplied promise returning function
* Strives for memory efficency - Automatic backpressure handling
* Line count support (not all xlsx files contain this metadata)
* Can transform column names
* Will reject on non xlsx data
* Battle tested in production code
* Skips empty rows (maybe it's a caveat)

## Caveats
* Only parses the first on sheet in a workbook
* Requires the first row to contain column names
* Skips empty rows (maybe it's a feature)

## Example Usage

```javascript
  const supersalmon = require('supersalmon')

  // The promise is resolved when the stream is completely processed
  // It resolves to the count of processed rows

  const processed = await supersalmon(

    // Any readable stream will work - remember that only the first sheet will be parsed
    createReadStream('huge.xlsx'),

    // process records one at a time - the argument object's keys are determined by col names
    ({name, surname}) => repository.insert({ name, surname }),

    // transform column names- column names will become the key names of the processed objects
    cols => colName => colName.toLowerCase().trim(),

    // the last function is called when the line count metadata is encountered in the stream
    lineCount => notifyLineCount(lineCount),
  );
```

Or see tests

## TODO List
[] Write some real documentation
[] Refactor
[] Add linter
[] Add more tests
[] Support files without column names in the first row
[] Parallel record processing
[] Port to typescript
[] Better API
[] Parse all sheets in a workbook

Issues and PRs more than welcome!