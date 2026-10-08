// Some stream dependencies use unprefixed util. Keep workbook contents out of diagnostic logs.
const util = require('node:util');
module.exports = {...util, debuglog: () => () => {}};
