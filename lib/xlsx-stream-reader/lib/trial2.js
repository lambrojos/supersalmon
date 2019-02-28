class XlsxStreamReaderWorkSheet {
  constructor(workBook, sheetName, workSheetId, workSheetStream) {

    this.id = workSheetId
    this.workBook = workBook
    this.name = sheetName
    this.options = workBook.options
    this.workSheetStream = workSheetStream
    this.rowCount = 0
    this.sheetData = {}
    this.inRows = false
    this.workingRow = {}
    this.currentCell = {}
    this.abortSheet = false
    this._handleWorkSheetStream()
  }
}
