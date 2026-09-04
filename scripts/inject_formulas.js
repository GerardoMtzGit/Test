const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, '..', 'data', 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx');
const wb = xlsx.readFile(filePath);

const rawVentas = xlsx.utils.sheet_to_json(wb.Sheets['raw_ventas']);
const depts = xlsx.utils.sheet_to_json(wb.Sheets['raw_departamento']);
const tiendas = xlsx.utils.sheet_to_json(wb.Sheets['raw_tiendas']);

// 1. Filtrar registros de 2012
const ventas2012 = rawVentas.filter(r => (r.semana_limpia || '').startsWith('2012'));
console.log('Filas 2012 para clean_ventas:', ventas2012.length);

// Diccionarios para valores iniciales (pre-renderizados para compatibilidad)
const deptMap = {};
depts.forEach(d => { deptMap[d.dept] = d.nombre_dept; });
const tiendaMap = {};
tiendas.forEach(t => { tiendaMap[t.tienda] = { tipo: t.tipo, tamano: Number(t.tamaño) || 0 }; });

// Construir clean_ventas con celdas que contienen fórmula Y valor precalculado
const cleanVentasRows = [
  ['tienda', 'tipo_tienda', 'tamano_m2', 'dept', 'nombre_dept', 'fecha', 'semana_limpia', 'ventas_semanales', 'esferiado']
];

ventas2012.forEach((r, idx) => {
  const rowNum = idx + 2;
  const tInfo = tiendaMap[r.tienda] || { tipo: 'N/A', tamano: 0 };
  const dName = deptMap[r.dept] || 'Depto 16 (Sin Asignar)';
  const vSales = Number(r.ventas_semanales) || 0;

  cleanVentasRows.push([
    r.tienda,
    { t: 's', f: `IFERROR(VLOOKUP(A${rowNum}, raw_tiendas!A:C, 2, FALSE), "N/A")`, v: tInfo.tipo },
    { t: 'n', f: `IFERROR(VLOOKUP(A${rowNum}, raw_tiendas!A:C, 3, FALSE), 0)`, v: tInfo.tamano },
    r.dept,
    { t: 's', f: `IFERROR(VLOOKUP(D${rowNum}, raw_departamento!A:B, 2, FALSE), "Depto 16 (Sin Asignar)")`, v: dName },
    r.fecha,
    r.semana_limpia,
    vSales,
    r.esferiado
  ]);
});

const wsClean = xlsx.utils.aoa_to_sheet(cleanVentasRows);
if (!wb.SheetNames.includes('clean_ventas')) {
  wb.SheetNames.splice(3, 0, 'clean_ventas');
}
wb.Sheets['clean_ventas'] = wsClean;

// 2. Construir PIVOT con fórmulas
const deptNamesList = [
  'Despensa y Básicos',
  'Comida Fresca',
  'Artículos del Hogar y Papel',
  'Salud y Bienestar',
  'Ropa',
  'Hogar y Temporada',
  'Cuidado del Bebé y la Familia',
  'Snacks y Bebidas',
  'Electrónica de Consumo',
  'Depto 16 (Sin Asignar)',
  'Automotriz',
  'Cuidado de Mascotas',
  'Juguetes y Juegos',
  'Oficina, Escuela y Manualidades',
  'Jardín y Vida al Aire Libre'
];

const pivotRows = [
  ['=== REPORTE 1: EFICIENCIA POR FORMATO DE TIENDA (VENTAS / M²) ==='],
  ['Tipo de Tienda', 'Ventas Totales 2012 (USD)', 'Área Total (m²)', 'Eficiencia ($/m²)'],
  [
    'Tipo A',
    { t: 'n', f: 'SUMIF(clean_ventas!B:B, "A", clean_ventas!H:H)', v: 335957905.20 },
    { t: 'n', f: 'SUMIF(raw_tiendas!B:B, "A", raw_tiendas!C:C)', v: 3899450 },
    { t: 'n', f: 'B3/C3', v: 86.16 }
  ],
  [
    'Tipo B',
    { t: 'n', f: 'SUMIF(clean_ventas!B:B, "B", clean_ventas!H:H)', v: 204378156.40 },
    { t: 'n', f: 'SUMIF(raw_tiendas!B:B, "B", raw_tiendas!C:C)', v: 1720242 },
    { t: 'n', f: 'B4/C4', v: 118.81 }
  ],
  [
    'Tipo C',
    { t: 'n', f: 'SUMIF(clean_ventas!B:B, "C", clean_ventas!H:H)', v: 18075458.98 },
    { t: 'n', f: 'SUMIF(raw_tiendas!B:B, "C", raw_tiendas!C:C)', v: 243250 },
    { t: 'n', f: 'B5/C5', v: 74.31 }
  ],
  [
    'Total Cadena',
    { t: 'n', f: 'SUM(B3:B5)', v: 558411520.58 },
    { t: 'n', f: 'SUM(C3:C5)', v: 5862942 },
    { t: 'n', f: 'B6/C6', v: 95.24 }
  ],
  [],
  ['=== REPORTE 2: PARTICIPACIÓN POR CATEGORÍA DE DEPARTAMENTO ==='],
  ['Departamento', 'Ventas Totales 2012 (USD)', 'Participación (%)']
];

deptNamesList.forEach((dName, i) => {
  const rowIdx = 10 + i;
  pivotRows.push([
    dName,
    { t: 'n', f: `SUMIF(clean_ventas!E:E, A${rowIdx}, clean_ventas!H:H)` },
    { t: 'n', f: `B${rowIdx}/$B$25` }
  ]);
});

pivotRows.push([
  'TOTAL VENTAS 2012',
  { t: 'n', f: 'SUM(B10:B24)', v: 558411520.58 },
  { t: 'n', f: 'SUM(C10:C24)', v: 1 }
]);

wb.Sheets['Pivot'] = xlsx.utils.aoa_to_sheet(pivotRows);

// 3. DASHBOARD con fórmulas
const dashRows = [
  ['DASHBOARD EJECUTIVO - WALMART VENTAS 2012'],
  ['Resumen para la Dirección Comercial (Ajustes de Presupuesto e Inventario)'],
  [],
  ['INDICADOR CLAVE (KPI)', 'VALOR CALCULADO', 'METODOLOGIA / FORMULA', 'INTERPRETACION ESTRATEGICA'],
  [
    'Ventas Totales 2012',
    { t: 'n', f: 'SUM(clean_ventas!H:H)', v: 558411520.58 },
    '=SUM(clean_ventas!H:H)',
    'Facturación neta agregada de las 45 sucursales en 2012.'
  ],
  [
    'Total Tiendas Operativas',
    { t: 'n', f: 'COUNTA(raw_tiendas!A2:A46)', v: 45 },
    '=COUNTA(raw_tiendas!A2:A46)',
    '22 tiendas Tipo A, 17 tiendas Tipo B y 6 tiendas Tipo C.'
  ],
  [
    'Formato Más Eficiente',
    'Tipo B ($118.81 / m²)',
    '=Pivot!D4',
    'Supera por +37.9% al Tipo A ($86.16) y +59.9% al Tipo C ($74.31).'
  ],
  [
    'Departamento Líder #1',
    { t: 's', f: 'Pivot!A10', v: 'Despensa y Básicos' },
    '=Pivot!A10',
    'Genera el 15.23% de las ventas globales ($85,052,985.88 USD).'
  ],
  [
    'Concentración Top 4 Departamentos',
    { t: 'n', f: 'SUM(Pivot!C10:C13)', v: 0.4556 },
    '=SUM(Pivot!C10:C13)',
    'Despensa, Frescos, Hogar y Salud concentran el 45.56% del ingreso.'
  ],
  [
    'Transacciones Semanales Limpias',
    { t: 'n', f: 'COUNT(clean_ventas!H:H)', v: 28845 },
    '=COUNT(clean_ventas!H:H)',
    'Volumen depurado sin pérdidas contables.'
  ]
];
wb.Sheets['Dashboard'] = xlsx.utils.aoa_to_sheet(dashRows);

// 4. RESUMEN C-F-I
const resumenRows = [
  ['Pregunta', 'KPI', 'Contexto', 'Insight / Hallazgo', 'Implicación Comercial'],
  [
    '¿Qué categorías de departamento fueron más eficientes para generar ventas en 2012?',
    'Ventas por m² (Ventas / Tamaño)',
    'Las 45 tiendas de Walmart operan en 3 formatos: Tipo A (Hipermercados, 177k m² prom.), Tipo B (Medianas, 101k m² prom.) y Tipo C (Pequeñas, 40k m² prom.).',
    'Las tiendas Tipo B alcanzaron la mayor eficiencia del negocio con $118.81/m², superando en un +37.9% a las Tipo A ($86.16/m²) y +59.9% a las Tipo C ($74.31/m²).',
    'Reasignar presupuesto de Capex priorizando aperturas y remodelaciones en formato Tipo B. Auditar el piso de venta en tiendas Tipo A para reducir metros ociosos y optimizar la densidad de pasillos.'
  ],
  [
    '¿Qué departamentos aportaron más al negocio y cuáles estuvieron por debajo de su potencial?',
    'Participación del Departamento (% s/ Total)',
    'En 2012 se vendieron $558.4M USD entre 14 departamentos catalogados más un departamento 16 no catalogado.',
    '4 departamentos concentran el 45.56% de la facturación: Despensa y Básicos (15.23%), Comida Fresca (10.66%), Artículos del Hogar y Papel (10.54%) y Salud y Bienestar (9.13%). Por el contrario, Jardinería (1.07%) y Oficina (1.47%) están rezagados.',
    'Blindar inventario y acuerdos de precio en Despensa y Frescos (productos gancho de alta rotación semanal). Reevaluar el espacio asignado a Jardinería y Papelería para cedérselo a categorías con mayor retorno.'
  ]
];
wb.Sheets['Resumen'] = xlsx.utils.aoa_to_sheet(resumenRows);

// 5. README QA Checks con fórmulas
const readmeAoa = xlsx.utils.sheet_to_json(wb.Sheets['README'], { header: 1 });
// Actualizar la sección de QA en el README
const updatedReadme = [
  [' 📊 Descripción del análisis Walmart 2012'],
  [],
  ['📑 Tablas Incluidas'],
  ['Hoja', 'Descripción', 'Tipo de hoja'],
  ['raw_ventas', 'Datos originales de ventas semanales, por tienda y departamento.', 'Raw Data'],
  ['raw_departamento', 'Catálogo de departamentos con sus nombres.', 'Lookup'],
  ['raw_tiendas', 'Catálogo de tiendas con tipo (A/B/C) y tamaño en m².', 'Lookup'],
  ['clean_ventas', 'Datos depurados de 2012 con uniones de catálogos, semana estandarizada y departamento.', 'Clean Data'],
  ['Pivot', 'Tablas dinámicas de Eficiencia ($/m²) y Participación por Departamento.', 'Pivot Data'],
  ['Dashboard', 'Panel ejecutivo de KPIs y métricas clave para la Dirección Comercial.', 'Dashboard'],
  ['Resumen', 'Matriz estratégica C-F-I (Context, Finding, Implication) para toma de decisiones.', 'Executive Summary'],
  [],
  ['📊 KPIs y Fórmulas'],
  ['KPI', 'Descripción', 'Fórmula en Sheets', 'Interpretación'],
  ['Ventas por m²', 'Ventas ajustadas por tamaño de tienda', '=SUMIF(clean_ventas!B:B, "B", clean_ventas!H:H) / SUMIF(raw_tiendas!B:B, "B", raw_tiendas!C:C)', 'Entre más alta la cifra, mayor eficiencia de ventas por metro cuadrado construido.'],
  ['Porcentaje de Participación', 'Participación relativa de cada departamento en la facturación', '=SUMIF(clean_ventas!E:E, "Despensa y Básicos", clean_ventas!H:H) / SUM(clean_ventas!H:H)', 'Identifica qué categorías impulsan el volumen y cuáles están por debajo de su potencial.'],
  [],
  ['✅ QA / Validaciones'],
  ['Chequeo de Calidad', 'Criterio de Verificación', 'Fórmula Aplicada en Sheets', 'Resultado'],
  ['¿Existen departamentos sin catalogar?', 'Detectar Depto 16 sin nombre en catálogo', '=COUNTIF(clean_ventas!E:E, "Depto 16 (Sin Asignar)")', '6,435 transacciones identificadas ($29.4M USD). Resuelto sin pérdida contable.'],
  ['¿Existen ventas negativas o devoluciones?', 'Buscar valores < 0', '=COUNTIF(clean_ventas!H:H, "<0")', '10 registros identificados (mínimo -$771.90). Mantenidos como devoluciones netas.'],
  ['¿Metros cuadrados con valor 0 o vacíos?', 'Revisar tamaño <= 0', '=COUNTIF(raw_tiendas!C:C, "<=0")', '0 errores. Las 45 tiendas tienen metros cuadrados válidos.']
];
wb.Sheets['README'] = xlsx.utils.aoa_to_sheet(updatedReadme);

xlsx.writeFile(wb, filePath);
console.log('✅ Archivo Excel reescrito con TODAS las fórmulas nativas de Google Sheets / Excel!');
