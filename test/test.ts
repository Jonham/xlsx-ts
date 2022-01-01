import assert from 'assert';
import { resolve } from 'path';
import { xlsx } from '../src/index';
import { CHAR_CHECK, __dirname } from './consts';

// const outputFolder = resolve('./test', './test-dist');
const outputFolder = resolve(__dirname, './test-dist');
const parseRoot = (f: string) => resolve(outputFolder, f);

function test(): void {
  const workbook = xlsx.createWorkbook(outputFolder, 'output.xlsx');
  const sheetName = 'test';
  const sheet = workbook.createSheet(sheetName, 10, 10);
  const colName = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (let row = 1; row <= 10; row++) {
    for (let col = 1; col <= 10; col++) {
      const v = `${colName[col]}-${row}`;
      sheet.set(col, row, v);

      if (row === 1) {
        sheet.set(col, row, {
          fill: {
            bgColor: '#ffff00',
            fgColor: '#333333',
          },
        });
      }
    }
  }

  // console.log('A1', sheet.data[1][1]);
  workbook.save((err) => {
    if (err) console.log(err);
  });
  const expected = '';
  // assert.equal(howLongTillLunch(...lunchtime), expected);
  console.log(`${CHAR_CHECK} ${expected}`);
}

test();

// let lunchtime = [12, 30];
// test(11, 30, 0, '1 hour');
// test(10, 30, 0, '2 hours');
// test(12, 25, 0, '5 minutes');
// test(12, 29, 15, '45 seconds');
// test(13, 30, 0, '23 hours');

// some of us like an early lunch
// lunchtime = [11, 0];
// test(10, 30, 0, '30 minutes');
