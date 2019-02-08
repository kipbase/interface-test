const axios = require('axios');
const XLSX = require('xlsx');

const config = require('./config.json');

/**
 * 生成表格数据指针
 * @param {string} column - 列头
 * @param {number} row - 行头
 */
function makePointer(column, row) {
  return column + row;
}

/**
 * 去除脏字符串中的字母，保留数字，转为数字类型
 * @param {string} str - 脏字符串
 */
function cleanLetter(str) {
  let reg = /[a-zA-Z]+/g;
  let result = str.replace(reg, '');
  return parseInt(result)
}

/**
 * 去除脏字符串中的数字，保留字母
 * @param {string} str - 脏字符串
 */
function cleanNumber(str) {
  let reg = /[0-9]+/g;
  let result = str.replace(reg, '');
  return result
}

let fileUri = config.file_name.indexOf('./') > 0 ?
              config.file_name :
              `./${config.file_name}`;

let ruleFile;

try {
  ruleFile = XLSX.readFile(fileUri);
} catch(e) {
  console.log('文件读取失败', e);
  return 0;
}

let ruleSheet = ruleFile.Sheets.Sheet1;
let pointer;
let valuePointer;
let interfaceArray = [];
let tempArray = [];
let initRow = 1;
let interfaceColumn = 'A';
let requestColumn = 'B';
let responseColumn = 'D';
let maxRow = cleanLetter(ruleSheet['!ref'].split(':').pop());
let maxColumn = cleanNumber(ruleSheet['!ref'].split(':').pop());

console.log(ruleSheet)

for (let i = initRow; i <= maxRow; i++) {
  pointer = makePointer(interfaceColumn, i);
  if (ruleSheet[pointer]) {
    if (tempArray.length === 0) {
      tempArray.push(i);
    } else {
      tempArray.push(i - 2);
      interfaceArray.push(tempArray);
      tempArray = [];
      tempArray.push(i);
    }
  }
}
tempArray.push(maxRow);
interfaceArray.push(tempArray);
tempArray = [];

for (let inter of interfaceArray) {
  let request = {}
  pointer = makePointer(interfaceColumn, inter[0]);
  let url = config.base_url + ruleSheet[pointer].v;
  for (let i = inter[0]; i <= inter[1]; i++) {
    pointer = makePointer(requestColumn, i)
    valuePointer = makePointer(String.fromCharCode(requestColumn.charCodeAt() + 1), i)
    if (ruleSheet[pointer]) {
      request[ruleSheet[pointer].v] = ruleSheet[valuePointer].v
    }
  }
  axios.get(url, request)
    .then(res => {
      // console.log(res.data)
    })
    .catch(e => {
      console.log(e)
    })
}