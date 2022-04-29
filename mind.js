const xlsx = require('node-xlsx');
const fs = require('fs');
const { Workbook, Topic, Marker, Zipper } = require('xmind');

/**
 * 源数据
 */
let data = fs.readFileSync('source/sourceExcel.xlsx');
const workSheetsFromFile = xlsx.parse(data);
const sheetData = workSheetsFromFile[0].data;
/**
 * 根数据
 */
const root = sheetData[0][0];

const [workbook, marker] = [new Workbook(), new Marker()];

let topic = new Topic({sheet: workbook.createSheet('sheet title', root)});
const zip = new Zipper({path: 'result', workbook, filename: 'target'});

const map = new Map();

sheetData.forEach((row, index) => {
  row.forEach((cell, cellIndex) => {
    if (cellIndex === 1) {
      if (map.has(cell + cellIndex)) {
        return;
      }
      topic = topic.on(topic.cid(root)).add({title: cell});
      map.set(cell + cellIndex, 1);
    } else if(cellIndex > 1) {
      if (map.has(cell + cellIndex)) {
        return;
      }
      topic = topic.on(topic.cid(row[cellIndex - 1])).add({title: cell});
      map.set(cell + cellIndex, 1);
    }
  });
});

zip.save().then(status => status && console.log('Saved.'));
