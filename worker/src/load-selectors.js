const fs = require('fs');
const path = require('path');

const selectorsPath = path.join(__dirname, 'selectors.json');
let _selectors = null;

function getSelectors(carrier) {
  if (!_selectors) {
    _selectors = JSON.parse(fs.readFileSync(selectorsPath, 'utf-8'));
  }
  const s = _selectors[carrier];
  if (!s) {
    throw new Error(`未対応の配送業者: ${carrier}`);
  }
  return s;
}

module.exports = { getSelectors };
