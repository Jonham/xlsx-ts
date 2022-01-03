// import assert from 'assert';
// import { resolve } from 'path';
import { xlsx } from '../dist/xlsx-ts.cjs.js';
import { CHAR_CHECK, __dirname } from './consts.mjs';

function test() {
  const workbook = xlsx.createWorkbook();
  const sheetName = 'test';
  const sheet = workbook.createSheet(sheetName, 10, 10);
  const colName = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (let row = 1; row <= 10; row++) {
    for (let col = 1; col <= 10; col++) {
      const v = `${colName[col]}-${row}`;
      sheet.set(col, row, v);

      if (row === 1) {
        sheet.set(col, row, {
          set: 'Red',
          font: {
            name: '宋体',
            sz: 11,
            color: 'FF0022FF',
            bold: true,
            iter: true,
            underline: true,
          },
          align: 'center',
          fill: {
            type: 'solid',
            fgColor: 'FFFF2200',
          },
        });
      }
    }
  }

  sheet.height(1, 20);

  workbook.save((err) => {
    if (err) console.log(err);
  });
  const expected = '';
  console.log(`${CHAR_CHECK} ${expected}`);
}

test();
