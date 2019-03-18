/*!
 * xlsx-stream-reader
 * Copyright(c) 2016 Brian Taber
 * MIT Licensed
 */

'use strict'

const Path = require('path')
const Fs = require('fs')
const Temp = require('temp')
const Stream = require('stream')
const unzipper = require('unzipper')

const StreamCreator = require(Path.join(__dirname, 'stream_creator'))
const WorkSheet = require(Path.join(__dirname, 'worksheet'))

module.exports = class WorkBook extends Stream {
  constructor(userOptions) {
    super()

    // default options
    this.defaults = {
      verbose: true,
      formatting: true,
      returnFormats: false
    }

    // overwrite default with userOptions
    this.options = Object.assign(this.defaults, userOptions)

    // WorkBook options
    this.workBookSharedStrings = []
    this.workBookInfo = {
      sheetRelationships: {},
      sheetRelationshipsNames: {}
    }
    this.parsedWorkBookInfo = false
    this.parsedWorkBookRels = false

    this.parsedSharedStrings = false
    this.waitingWorkSheets = []
    this.workBookStyles = []
    this.formatCodes = {}
    this.xfs = {}
    this.abortBook = false
    this._handleWorkBookStream()
  }

  write(){}
  end(){}

  _handleWorkBookStream() {
    var match
    this.on('pipe', srcPipe => {
      srcPipe.pipe(unzipper.Parse())
        .on('error', err => {
          this.emit('error', err)
        })
        .on('entry', entry => {
          if (this.abortBook) {
            entry.autodrain()
            return
          }
          switch (entry.path) {
          case 'xl/workbook.xml':
            this._parseXML(entry, this._parseWorkBookInfo, () => {
              this.parsedWorkBookInfo = true
              this.emit('workBookInfo')
            })
            break
          case 'xl/_rels/workbook.xml.rels':
            this._parseXML(entry, this._parseWorkBookRels, () => {
              this.parsedWorkBookRels = true
              this.emit('workBookRels')
            })
            break
          case '_rels/.rels':
            entry.autodrain()
            break
          case 'xl/sharedStrings.xml': {
            let currentString = ''

            this._parseXML(entry,
              nodeData => {
                var self = this
                if(nodeData[0] && nodeData[0].name === 'si') {
                  if(currentString !== null) {
                    this.workBookSharedStrings.push(currentString)
                  }
                  currentString = ''
                }

                var nodeObjValue = nodeData.pop()
                var nodeObjName = nodeData.pop()

                if (nodeObjName && nodeObjName.name === 't') {
                  currentString = `${currentString}${nodeObjValue}`
                } else if (nodeObjValue && typeof nodeObjValue === 'object' && nodeObjValue.hasOwnProperty('name') && nodeObjValue.name === 't') {
                  self.workBookSharedStrings.push('')
                }
              }
              , () => {
                this.workBookSharedStrings.push(currentString)
                this.parsedSharedStrings = true
                this.emit('sharedStrings')
              })
            break
          }
          case 'xl/styles.xml':
            this._parseXML(entry, this._parseStyles, () => {
              var cellXfsIndex = this.workBookStyles.findIndex( item => {
                return item.name === 'cellXfs'
              })
              this.xfs = this.workBookStyles.filter( (item, index) => {
                return item.name === 'xf' && index > cellXfsIndex
              })
              this.emit('styles')
            })
            break
          default:
            if ((match = entry.path.match(/xl\/(worksheets\/sheet(\d+)\.xml)/))) {
              var sheetPath = match[1]
              var sheetNo = match[2]

              if (this.parsedWorkBookInfo === false ||
                  this.parsedWorkBookRels === false ||
                  this.parsedSharedStrings === false ||
                  this.waitingWorkSheets.length > 0
              ) {
                var stream = Temp.createWriteStream()

                this.waitingWorkSheets.push({ sheetNo: sheetNo, name: entry.path, path: stream.path, sheetPath: sheetPath })

                entry.pipe(stream)
              } else {
                var name = this._getSheetName(sheetPath)
                var workSheet = new WorkSheet(this, name, sheetNo, entry)

                this.emit('worksheet', workSheet)
              }
            } else if ((match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)\.xml.rels/))) {
              entry.autodrain()
            } else {
              entry.autodrain()
            }
            break
          }
        })
        .on('close', () => {
          if (this.waitingWorkSheets.length > 0) {
            var currentBook = 0
            var processBooks = () => {
              var sheetInfo = this.waitingWorkSheets[currentBook]
              var workSheetStream = Fs.createReadStream(sheetInfo.path)
              var name = this._getSheetName(sheetInfo.sheetPath)
              var workSheet = new WorkSheet(this, name, sheetInfo.sheetNo, workSheetStream)

              workSheet.on('end', () => {
                ++currentBook
                if (currentBook === this.waitingWorkSheets.length) {
                  Temp.cleanupSync()
                  setImmediate(this.emit.bind(this), 'end')
                } else {
                  setImmediate(processBooks)
                }
              })

              setImmediate(this.emit.bind(this), 'worksheet', workSheet)
            }
            setImmediate(processBooks)
          } else {
            setImmediate(this.emit.bind(this), 'end')
          }
        })
    })
  }

  abort() {
    this.abortBook = true
  }

  _parseXML(entryStream, entryHandler, endHandler) {
    let isErred = false

    let tmpNode = []
    let tmpNodeEmit = false

    const parser = StreamCreator()

    entryStream.on('end', () => {
      if (this.abortBook) return
      if (!isErred) setImmediate(endHandler)
    })

    parser.on('error', error => {
      if (this.abortBook) return
      isErred = true

      this.emit('error', error)
    })

    parser.on('opentag', node => {
      if (node.name === 'rPh') {
        this.abortBook = true
        return
      }
      if (this.abortBook) return
      if (Object.keys(node.attributes).length === 0) {
        delete (node.attributes)
      }
      if (node.isSelfClosing) {
        if (tmpNode.length > 0) {
          entryHandler.call(this, tmpNode)
          tmpNode = []
        }
        tmpNodeEmit = true
      }
      delete (node.isSelfClosing)
      tmpNode.push(node)
    })

    parser.on('text', text => {
      if (this.abortBook) return
      tmpNodeEmit = true
      tmpNode.push(text)
    })

    parser.on('closetag', nodeName => {
      if (nodeName === 'rPh') {
        this.abortBook = false
        return
      }
      if (this.abortBook) return
      if (tmpNodeEmit) {
        entryHandler.call(this, tmpNode)
        tmpNodeEmit = false
        tmpNode = []
      } else if (tmpNode.length && tmpNode[tmpNode.length - 1] && tmpNode[tmpNode.length - 1].name === nodeName) {
        tmpNode.push('')
        entryHandler.call(this, tmpNode)
        tmpNodeEmit = false
        tmpNode = []
      }
      tmpNode.splice(-1, 1)
    })

    try {
      entryStream.pipe(parser)
    } catch (error) {
      this.emit('error', error)
    }
  }

  _getSharedString(stringIndex) {
    if (stringIndex > this.workBookSharedStrings.length) {
      if (this.options.verbose) {
        this.emit('error', 'missing shared string: ' + stringIndex)
      }
      return
    }
    return this.workBookSharedStrings[stringIndex]
  }

  _parseStyles(nodeData) {
    nodeData.forEach( data => {
      if (data.name === 'numFmt') {
        this.formatCodes[data.attributes.numFmtId] = data.attributes.formatCode
      }
      this.workBookStyles.push(data)
    })
  }

  _parseWorkBookInfo (nodeData) {
    nodeData.forEach( data => {
      if (data.name === 'sheet') {
        this.workBookInfo.sheetRelationshipsNames[data.attributes['r:id']] = data.attributes.name
      }
    })
  }

  _parseWorkBookRels(nodeData) {
    nodeData.forEach( data => {
      if (data.name === 'Relationship') {
        this.workBookInfo.sheetRelationships[data.attributes.Target] = data.attributes.Id
      }
    })
  }

  _getSheetName (sheetPath) {
    return this.workBookInfo.sheetRelationshipsNames[this.workBookInfo.sheetRelationships[sheetPath]]
  }
}