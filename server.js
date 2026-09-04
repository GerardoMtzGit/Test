const http = require('http');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = __dirname;
const EXCEL_PATH = path.join(__dirname, 'data', 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx');

let cachedWorkbook = null;
let summaryData = null;

function loadExcel() {
  try {
    if (fs.existsSync(EXCEL_PATH)) {
      console.log('Cargando archivo Excel en memoria...');
      cachedWorkbook = xlsx.readFile(EXCEL_PATH);
      console.log('Hojas disponibles:', cachedWorkbook.SheetNames);

      // Precalcular KPIs para respuesta instantánea
      const rawVentas = xlsx.utils.sheet_to_json(cachedWorkbook.Sheets['raw_ventas'] || {});
      const depts = xlsx.utils.sheet_to_json(cachedWorkbook.Sheets['raw_departamento'] || {});
      const tiendas = xlsx.utils.sheet_to_json(cachedWorkbook.Sheets['raw_tiendas'] || {});

      const deptMap = {};
      depts.forEach(d => { deptMap[d.dept] = d.nombre_dept; });

      let totalSales = 0;
      const deptSales = {};

      rawVentas.forEach(r => {
        const val = Number(r.ventas_semanales) || 0;
        totalSales += val;
        const deptName = deptMap[r.dept] || `Depto ${r.dept}`;
        deptSales[deptName] = (deptSales[deptName] || 0) + val;
      });

      const topDepts = Object.entries(deptSales)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([name, sales]) => ({ name, sales }));

      summaryData = {
        fileName: 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx',
        sheetNames: cachedWorkbook.SheetNames,
        totalSales,
        formattedSales: totalSales.toLocaleString('en-US', { style: 'currency', currency: 'USD' }),
        totalTransactions: rawVentas.length,
        totalStores: tiendas.length,
        totalDepartments: depts.length,
        topDepartments: topDepts
      };
      console.log('✅ Datos de Walmart procesados y listos.');
    }
  } catch (err) {
    console.error('Error cargando Excel:', err);
  }
}

loadExcel();

const MIME_TYPES = {
  '.html': 'text/html; charset=UTF-8',
  '.css': 'text/css; charset=UTF-8',
  '.js': 'application/javascript; charset=UTF-8',
  '.json': 'application/json',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer((req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = parsedUrl.pathname;

  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  // API Healthcheck
  if (pathname === '/api/healthcheck') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({
      status: 'ok',
      service: 'VibeCode Local Server',
      uptime: process.uptime(),
      timestamp: new Date().toISOString()
    }));
  }

  // API Excel Summary
  if (pathname === '/api/excel/summary') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify(summaryData || { error: 'No data loaded' }));
  }

  // API Excel Sheet data
  if (pathname === '/api/excel/sheet') {
    if (!cachedWorkbook) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ error: 'Workbook not loaded' }));
    }

    const sheetName = parsedUrl.searchParams.get('name') || cachedWorkbook.SheetNames[0];
    const page = Math.max(1, parseInt(parsedUrl.searchParams.get('page') || '1', 10));
    const limit = Math.min(200, Math.max(10, parseInt(parsedUrl.searchParams.get('limit') || '50', 10)));
    const search = (parsedUrl.searchParams.get('search') || '').toLowerCase().trim();

    const sheet = cachedWorkbook.Sheets[sheetName];
    if (!sheet) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      return res.end(JSON.stringify({ error: `Sheet '${sheetName}' not found` }));
    }

    // Convert to row arrays
    const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const headers = rawRows[0] || [];
    let dataRows = rawRows.slice(1);

    if (search) {
      dataRows = dataRows.filter(row =>
        row.some(cell => String(cell).toLowerCase().includes(search))
      );
    }

    const totalRows = dataRows.length;
    const totalPages = Math.ceil(totalRows / limit) || 1;
    const startIndex = (page - 1) * limit;
    const paginatedRows = dataRows.slice(startIndex, startIndex + limit);

    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({
      sheetName,
      headers,
      rows: paginatedRows,
      page,
      limit,
      totalRows,
      totalPages
    }));
  }

  // Download raw Excel
  if (pathname === '/api/excel/download') {
    if (!fs.existsSync(EXCEL_PATH)) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      return res.end('Archivo no encontrado');
    }
    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Proyecto_2_Resumen_Ejecutivo_Walmart.xlsx"'
    });
    return fs.createReadStream(EXCEL_PATH).pipe(res);
  }

  // Static files
  let safePath = path.normalize(pathname).replace(/^(\.\.[\/\\])+/, '');
  if (safePath === '/' || safePath === '\\') {
    safePath = '/index.html';
  }

  const filePath = path.join(PUBLIC_DIR, safePath);

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=UTF-8' });
      return res.end(`404 Not Found: ${safePath}`);
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  });
});

server.listen(PORT, () => {
  console.log(`========================================`);
  console.log(`🚀 VibeCode Server corriendo en:`);
  console.log(`👉 http://localhost:${PORT}`);
  console.log(`========================================`);
});
