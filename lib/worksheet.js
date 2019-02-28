/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict'

const ssf = require('ssf')
const Stream = require('stream')

module.exports = class WorkSheet extends Stream {
  constructor(workBook, sheetName, workSheetId, workSheetStream) {
    super()

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

  write(){}
  end(){}

  _handleWorkSheetStream() {
    this.on('pipe', srcPipe => {
      this.workBook._parseXML.call(this, srcPipe, this._handleWorkSheetNode, () => {
        if (this.workingRow.name) {
          delete (this.workingRow.name)
          this.emit('row', this.workingRow)
          this.workingRow = {}
        }
        this.emit('end')
      })
    })
  }

  getColumnNumber(columnName) {
    let i = columnName.search(/\d/)
    let colNum = 0
    columnName = +columnName.replace(/\D/g, letter => {
      colNum += (parseInt(letter, 36) - 9) * Math.pow(26, --i)
      return ''
    })
    return colNum
  }

  getColumnName(columnNumber) {
    let columnName = ''
    let dividend = parseInt(columnNumber)
    let modulo = 0
    while (dividend > 0) {
      modulo = (dividend - 1) % 26
      columnName = String.fromCharCode(65 + modulo).toString() + columnName
      dividend = Math.floor(((dividend - modulo) / 26))
    }
    return columnName
  }

  process() {
    this.workSheetStream.pipe(this)
  }

  abort() {
    this.abortSheet = true
  }

  _getNumberFormatId(workingCell, workingVal) {
    if (!(this.options.formatting && workingVal)) {
      return null
    }
    return workingCell.attributes.s ? this.workBook.xfs[workingCell.attributes.s].attributes.numFmtId : null 
  }

  _handleWorkSheetNode(nodeData) {

    if (this.abortSheet) {
      return
    }

    this.sheetData['cols'] = []

    switch (nodeData[0].name) {
    case 'worksheet':
    case 'sheetPr':
    case 'pageSetUpPr':
      return

    case 'printOptions':
    case 'pageMargins':
    case 'pageSetup':
      this.inRows = false
      if (this.workingRow.name) {
        delete (this.workingRow.name)
        this.emit('row', this.workingRow)
        this.workingRow = {}
      }
      break

    case 'cols':
      return

    case 'col':
      delete (nodeData[0].name)
      this.sheetData['cols'].push(nodeData[0])
      return

    case 'sheetData':
      this.inRows = true

      nodeData.shift()

    case 'row': // eslint-disable-line no-fallthrough
      if (this.workingRow.name) {
        delete (this.workingRow.name)
        this.emit('row', this.workingRow)
        this.workingRow = {}
      }

      ++this.rowCount

      this.workingRow = nodeData.shift() || {}
      if (typeof this.workingRow !== 'object') {
        this.workingRow = {}
      }
      this.workingRow.values = []
      this.workingRow.formulas = []
      if (this.options.returnFormats) {
        this.workingRow.formats = []
      }
      break
    }

    if (this.inRows === true) {
      let workingCell = nodeData.shift()
      const workingPart = nodeData.shift()
      let workingVal = nodeData.shift()

      if (!workingCell) {
        return
      }

      if (workingCell && workingCell.attributes && workingCell.attributes.r) {
        this.currentCell = workingCell
      }

      if (workingCell.name === 'c') {
        const cellNum = this.getColumnNumber(workingCell.attributes.r)

        if (workingPart && workingPart.name && workingPart.name === 'f') {
          this.workingRow.formulas[cellNum] = workingVal
        }

        // ST_CellType
        switch (workingCell.attributes.t) {
        case 's':
        // shared string
          var index = parseInt(workingVal)
          workingVal = this.workBook._getSharedString(index)

          this.workingRow.values[cellNum] = workingVal || ''

          workingCell = {}
          break
        case 'inlineStr':
        // inline string
          this.workingRow.values[cellNum] = nodeData.shift() || ''

          workingCell = {}
          break
        case 'str':
        // string (formula)
        case 'b': // eslint-disable-line no-fallthrough
        // boolean
        case 'n': // eslint-disable-line no-fallthrough
        // number
        case 'e': // eslint-disable-line no-fallthrough
        // error
        default: // eslint-disable-line no-fallthrough
          var formatId = this._getNumberFormatId(workingCell, workingVal)
          if (formatId !== null) {
            var format = this.workBook.formatCodes[formatId]
            if (typeof format === 'undefined') {
              try {
                workingVal = ssf.format(Number(formatId), Number(workingVal))
              } catch (e) {
                workingVal = ''
              }
            } else if (format !== 'General') {
              try {
                workingVal = ssf.format(format, Number(workingVal))
              } catch (e) {
                workingVal = ''
              }
            }
          } else if (!isNaN(parseFloat(workingVal))) { // this is number
            workingVal = parseFloat(parseFloat(workingVal)) // parse to float or int
          }
          this.workingRow.values[cellNum] = workingVal || ''
          if (this.options.returnFormats) {
            this.workingRow.formats[cellNum] = format || ssf.get_table()[formatId]
          }
          workingCell = {}
        }
      }
      if (workingCell.name === 'v') {
        var colNum = this.getColumnNumber(this.currentCell.attributes.r)

        this.currentCell = {}

        this.workingRow.values[colNum] = workingPart || ''
      }
    } else {
      if (this.sheetData[nodeData[0].name]) {
        if (!Array.isArray(this.sheetData[nodeData[0].name])) {
          this.sheetData[nodeData[0].name] = [this.sheetData[nodeData[0].name]]
        }
        this.sheetData[nodeData[0].name].push(nodeData)
      } else {
        if (nodeData[0].name) {
          this.sheetData[nodeData[0].name] = nodeData
        }
      }
    }
  }
}