function runExport() {
  console.log('runExport');

  const { xlsx } = XLSXts;
  const workbook = xlsx.createWorkbook('', 'output.xlsx');
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

  workbook.generate((err, zip) => {
    // let buffer;
    let args = { type: 'blob', mimeType: 'application/vnd.ms-excel;' };
    args.compressed = 'DEFLATE';

    zip.generateAsync(args).then((blob) => {
      if (err) {
        console.error(err);
        return;
      }

      const filename = 'test.xlsx';
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', filename);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    });
  });
}
