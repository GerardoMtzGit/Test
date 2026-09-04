const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

async function check() {
  const gitPath = 'C:\\Program Files\\Git\\cmd\\git.exe';
  const buf = execSync(`"${gitPath}" show "a0c86f1:data/Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx"`, { maxBuffer: 20 * 1024 * 1024 });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const ws = wb.getWorksheet('Dashboard');
  console.log('Original Dashboard rows:');
  ws.eachRow({ includeEmpty: true }, (row, rNum) => {
    row.eachCell({ includeEmpty: false }, (cell, cNum) => {
      console.log(`Cell ${cell.address} (R${rNum}C${cNum}): ${JSON.stringify(cell.value)}`);
      if (cell.border) {
        console.log(`   Border on ${cell.address}:`, JSON.stringify(cell.border));
      }
    });
  });

  console.log('Original Merged ranges:');
  console.log(ws._merges ? Object.keys(ws._merges) : 'none');
}

check().catch(console.error);
