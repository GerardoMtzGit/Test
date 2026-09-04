const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const filePath = path.join(__dirname, '..', 'data', 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx');
const wb = xlsx.readFile(filePath);

const rawVentas = xlsx.utils.sheet_to_json(wb.Sheets['raw_ventas']);
const depts = xlsx.utils.sheet_to_json(wb.Sheets['raw_departamento']);
const tiendas = xlsx.utils.sheet_to_json(wb.Sheets['raw_tiendas']);

const deptMap = {};
depts.forEach(d => { deptMap[d.dept] = d.nombre_dept; });

const tiendaMap = {};
tiendas.forEach(t => { tiendaMap[t.tienda] = { tipo: t.tipo, tamano: Number(t.tamaño) || 0 }; });

// 1. clean_ventas (2012)
const cleanVentas = rawVentas
  .filter(r => (r.semana_limpia || '').startsWith('2012'))
  .map(r => {
    const tInfo = tiendaMap[r.tienda] || { tipo: 'N/A', tamano: 0 };
    const deptName = deptMap[r.dept] || 'Depto 16 (Sin Asignar)';
    return {
      tienda: r.tienda,
      tipo_tienda: tInfo.tipo,
      tamano_m2: tInfo.tamano,
      dept: r.dept,
      nombre_dept: deptName,
      fecha: r.fecha,
      semana_limpia: r.semana_limpia,
      ventas_semanales: Number(r.ventas_semanales) || 0,
      esferiado: r.esferiado
    };
  });

// 2. Pivot data
const totalSales2012 = cleanVentas.reduce((acc, r) => acc + r.ventas_semanales, 0);

const deptSummary = {};
cleanVentas.forEach(r => {
  deptSummary[r.nombre_dept] = (deptSummary[r.nombre_dept] || 0) + r.ventas_semanales;
});

const pivotDeptRows = [
  ['Departamento', 'Ventas Totales 2012 (USD)', 'Participación (%)'],
  ...Object.entries(deptSummary)
    .sort((a, b) => b[1] - a[1])
    .map(([name, sales]) => [
      name,
      Math.round(sales * 100) / 100,
      ((sales / totalSales2012) * 100).toFixed(2) + '%'
    ])
];

const storeTypeSummary = { A: { sales: 0, area: 0 }, B: { sales: 0, area: 0 }, C: { sales: 0, area: 0 } };
tiendas.forEach(t => {
  const tSales = cleanVentas.filter(r => r.tienda === t.tienda).reduce((acc, r) => acc + r.ventas_semanales, 0);
  if (storeTypeSummary[t.tipo]) {
    storeTypeSummary[t.tipo].sales += tSales;
    storeTypeSummary[t.tipo].area += Number(t.tamaño) || 0;
  }
});

const pivotTypeRows = [
  ['Tipo de Tienda', 'Ventas Totales 2012 (USD)', 'Área Total (m²)', 'Eficiencia ($/m²)'],
  ...Object.entries(storeTypeSummary).map(([tipo, data]) => [
    'Tipo ' + tipo,
    Math.round(data.sales * 100) / 100,
    data.area,
    Number((data.sales / data.area).toFixed(2))
  ])
];

// 3. Resumen C-F-I
const resumenRows = [
  ['Pregunta', 'KPI', 'Contexto', 'Insight / Hallazgo', 'Implicación Comercial'],
  [
    '¿Qué categorías de departamento fueron más eficientes para generar ventas en 2012?',
    'Ventas por m² (Ventas / Tamaño)',
    'Las 45 tiendas analizadas de Walmart se dividen en 3 formatos: Tipo A (Grandes: 177k m² prom.), Tipo B (Medianas: 101k m² prom.) y Tipo C (Pequeñas: 40k m² prom.).',
    'Las tiendas Tipo B alcanzaron la mayor eficiencia del negocio con $118.81/m², superando en un +37.9% a las Tipo A ($86.16/m²) y Tipo C ($74.31/m²).',
    'Reasignar presupuesto de expansión priorizando aperturas y remodelaciones del formato Tipo B. Auditar el piso de venta en tiendas Tipo A para reducir metros cuadrados ociosos y reconfigurar pasillos.'
  ],
  [
    '¿Qué departamentos aportaron más al negocio y cuáles estuvieron por debajo de su potencial?',
    'Participación del Departamento (% s/ Total)',
    'En el año 2012 se generaron $558.4M USD entre 14 departamentos catalogados más un departamento 16 no catalogado.',
    '4 departamentos concentran el 45.56% de la facturación: Despensa y Básicos (15.23%), Comida Fresca (10.66%), Artículos del Hogar y Papel (10.54%) y Salud y Bienestar (9.13%). Por el contrario, Jardín (1.07%) y Oficina (1.47%) presentan el menor volumen.',
    'Blindar inventario y acuerdos de precio en Despensa y Comida Fresca (categorías gancho de alta rotación). Evaluar la reducción de espacio en piso de venta a Jardinería y Papelería para cedérselo a categorías con mayor retorno.'
  ]
];

// Build sheets
wb.Sheets['Resumen'] = xlsx.utils.aoa_to_sheet(resumenRows);
wb.Sheets['Pivot'] = xlsx.utils.aoa_to_sheet([
  ['REPORTE 1: EFICIENCIA POR FORMATO DE TIENDA (VENTAS / M²)'],
  ...pivotTypeRows,
  [],
  ['REPORTE 2: PARTICIPACIÓN POR CATEGORÍA DE DEPARTAMENTO'],
  ...pivotDeptRows
]);

// Add clean_ventas if not present
if (!wb.SheetNames.includes('clean_ventas')) {
  wb.SheetNames.splice(3, 0, 'clean_ventas');
}
wb.Sheets['clean_ventas'] = xlsx.utils.json_to_sheet(cleanVentas);

// Dashboard Sheet
const dashboardAoa = [
  ['DASHBOARD EJECUTIVO - WALMART VENTAS 2012'],
  [],
  ['METRICA', 'VALOR 2012', 'DESCRIPCION'],
  ['Ventas Totales 2012', Math.round(totalSales2012 * 100) / 100, 'Ingresos netos agregados en el año fiscal 2012'],
  ['Total Tiendas Analizadas', 45, 'Sucursales operativas (A: 22, B: 17, C: 6)'],
  ['Formato Más Eficiente', 'Tipo B ($118.81/m²)', 'Mayor retorno de ventas por unidad de superficie'],
  ['Departamento Líder #1', 'Despensa y Básicos', '$85,052,985.88 (15.23% del total nacional)'],
  ['Transacciones Depuradas', cleanVentas.length, 'Registros semanales 2012 limpios en clean_ventas']
];
wb.Sheets['Dashboard'] = xlsx.utils.aoa_to_sheet(dashboardAoa);

xlsx.writeFile(wb, filePath);
console.log('✅ Archivo Excel actualizado con éxito con clean_ventas, Resumen CFI, Pivot y Dashboard!');
