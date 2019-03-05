'use strict'

const Sax = require('sax')

module.exports = StreamCreator

function StreamCreator(userStrict, userOptions) {

  // default value
  const saxStrict = true
  const saxOptions = {
    trim: true,
    normalize: true,
    position: true,
    strictEntities: true
  }

  // overwrite with user options
  const strict = Object.assign(saxStrict, userStrict)
  const options = Object.assign(saxOptions, userOptions)
  
  return Sax.createStream(strict, options)
}