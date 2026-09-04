const ExcelJS = require('exceljs');
const path = require('path');

async function buildDashboard() {
  const filePath = path.join(__dirname, '..', 'data', 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  let wsDash = wb.getWorksheet('Dashboard');
  if (!wsDash) {
    wsDash = wb.addWorksheet('Dashboard');
  } else {
    // Limpiar filas anteriores de Dashboard
    while (wsDash.rowCount > 0) {
      wsDash.spliceRows(1, 1);
    }
  }

  // Configuración de anchos de columna
  wsDash.columns = [
    { width: 4 },  // A: Margen
    { width: 14 }, // B
    { width: 16 }, // C: KPI 1
    { width: 16 }, // D: KPI 1
    { width: 15 }, // E: KPI 1
    { width: 18 }, // F: KPI Central
    { width: 18 }, // G: KPI Central
    { width: 24 }, // H: KPI 2 / Depto
    { width: 18 }, // I: KPI 2
    { width: 16 }, // J: KPI 2
    { width: 16 }, // K: KPI 2
    { width: 22 }, // L: Sparkline / Barra
    { width: 6 }   // M
  ];

  // Paleta de Estilos
  const redHero = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFF0000' } };
  const redLarge = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFF0000' } };
  const redMedium = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFCC0000' } };
  const headerDark = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const textMuted = { name: 'Calibri', size: 9, italic: true, color: { argb: 'FF64748B' } };
  const textFormula = { name: 'Consolas', size: 9, bold: true, color: { argb: 'FF0284C7' } };

  const cardHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  const cardBodyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } };
  const tableHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  const rowZebraFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };

  const boxBorder = {
    top: { style: 'medium', color: { argb: 'FF1E293B' } },
    left: { style: 'medium', color: { argb: 'FF1E293B' } },
    bottom: { style: 'medium', color: { argb: 'FF1E293B' } },
    right: { style: 'medium', color: { argb: 'FF1E293B' } }
  };

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
  };

  // Título Superior del Dashboard
  wsDash.getRow(2).height = 24;
  wsDash.getCell('B2').value = 'DASHBOARD EJECUTIVO - WALMART VENTAS 2012';
  wsDash.getCell('B2').font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FF0F172A' } };

  wsDash.getRow(3).height = 18;
  wsDash.getCell('B3').value = 'Dirección Comercial · Panel de Toma de Decisiones de Presupuesto e Inventario';
  wsDash.getCell('B3').font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF64748B' } };

  // ==========================================
  // TARJETA 1: KPI 1 - EFICIENCIA POR M² (C6:E9)
  // Reemplaza el recuadro original "espacio para kpi"
  // ==========================================
  wsDash.mergeCells('C6:E6');
  const kpi1Header = wsDash.getCell('C6');
  kpi1Header.value = 'KPI 1: EFICIENCIA POR METRO CUADRADO';
  kpi1Header.font = headerDark;
  kpi1Header.fill = cardHeaderFill;
  kpi1Header.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('C7:E8');
  const kpi1Val = wsDash.getCell('C7');
  kpi1Val.value = '$118.81 / m²';
  kpi1Val.font = redHero;
  kpi1Val.fill = cardBodyFill;
  kpi1Val.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('C9:E9');
  const kpi1Sub = wsDash.getCell('C9');
  kpi1Sub.value = '🏆 Formato Ganador: Tipo B (+37.9% vs Tipo A)';
  kpi1Sub.font = redMedium;
  kpi1Sub.fill = cardBodyFill;
  kpi1Sub.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('C10:E10');
  const kpi1Form = wsDash.getCell('C10');
  kpi1Form.value = 'Fórmula: =SUMAR.SI(clean_ventas!B:B,"B",clean_ventas!H:H)/SUMAR.SI(raw_tiendas!B:B,"B",raw_tiendas!C:C)';
  kpi1Form.font = textFormula;
  kpi1Form.alignment = { vertical: 'middle', horizontal: 'center' };

  // Aplicar bordes al Box 1
  ['C6','D6','E6','C7','D7','E7','C8','D8','E8','C9','D9','E9'].forEach(addr => {
    wsDash.getCell(addr).border = boxBorder;
  });

  // ==========================================
  // TARJETA CENTRAL: VENTAS TOTALES 2012 (F6:H9)
  // ==========================================
  wsDash.mergeCells('F6:H6');
  const kpiTotHeader = wsDash.getCell('F6');
  kpiTotHeader.value = 'VENTAS TOTALES NETAS 2012';
  kpiTotHeader.font = headerDark;
  kpiTotHeader.fill = cardHeaderFill;
  kpiTotHeader.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('F7:H8');
  const kpiTotVal = wsDash.getCell('F7');
  kpiTotVal.value = '$558,411,520.58 USD';
  kpiTotVal.font = redHero;
  kpiTotVal.fill = cardBodyFill;
  kpiTotVal.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('F9:H9');
  const kpiTotSub = wsDash.getCell('F9');
  kpiTotSub.value = '45 Tiendas Analizadas · 28,845 Semanas Limpias';
  kpiTotSub.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF334155' } };
  kpiTotSub.fill = cardBodyFill;
  kpiTotSub.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('F10:H10');
  const kpiTotForm = wsDash.getCell('F10');
  kpiTotForm.value = 'Fórmula: =SUMA(clean_ventas!H:H)';
  kpiTotForm.font = textFormula;
  kpiTotForm.alignment = { vertical: 'middle', horizontal: 'center' };

  ['F6','G6','H6','F7','G7','H7','F8','G8','H8','F9','G9','H9'].forEach(addr => {
    wsDash.getCell(addr).border = boxBorder;
  });

  // ==========================================
  // TARJETA 2: KPI 2 - PARTICIPACIÓN DEPARTAMENTO (I6:K9)
  // Reemplaza el recuadro original "espacio para kpi"
  // ==========================================
  wsDash.mergeCells('I6:K6');
  const kpi2Header = wsDash.getCell('I6');
  kpi2Header.value = 'KPI 2: PARTICIPACION POR DEPARTAMENTO';
  kpi2Header.font = headerDark;
  kpi2Header.fill = cardHeaderFill;
  kpi2Header.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('I7:K8');
  const kpi2Val = wsDash.getCell('I7');
  kpi2Val.value = '15.23% · Despensa y Básicos';
  kpi2Val.font = redHero;
  kpi2Val.fill = cardBodyFill;
  kpi2Val.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('I9:K9');
  const kpi2Sub = wsDash.getCell('I9');
  kpi2Sub.value = '⭐ Depto Líder ($85,052,985.88 USD Facturados)';
  kpi2Sub.font = redMedium;
  kpi2Sub.fill = cardBodyFill;
  kpi2Sub.alignment = { vertical: 'middle', horizontal: 'center' };

  wsDash.mergeCells('I10:K10');
  const kpi2Form = wsDash.getCell('I10');
  kpi2Form.value = 'Fórmula: =SUMAR.SI(clean_ventas!E:E,"Despensa y Básicos",clean_ventas!H:H)/SUMA(clean_ventas!H:H)';
  kpi2Form.font = textFormula;
  kpi2Form.alignment = { vertical: 'middle', horizontal: 'center' };

  ['I6','J6','K6','I7','J7','K7','I8','J8','K8','I9','J9','K9'].forEach(addr => {
    wsDash.getCell(addr).border = boxBorder;
  });

  // ==========================================
  // SECCIÓN: ESPACIO PARA GRÁFICOS (Fila 13 en adelante)
  // ==========================================

  // Título Gráfico 1 (Izquierda)
  wsDash.mergeCells('B13:E13');
  const g1Title = wsDash.getCell('B13');
  g1Title.value = '📊 GRÁFICO 1: COMPARATIVA DE EFICIENCIA POR FORMATO DE TIENDA ($/m²)';
  g1Title.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  g1Title.fill = tableHeaderFill;
  g1Title.alignment = { vertical: 'middle', horizontal: 'left' };

  // Encabezados Tabla Gráfico 1
  const g1Headers = ['Formato', 'Ventas 2012', 'Superficie', 'Eficiencia ($/m²)', 'Gráfico Visual'];
  ['B14','C14','D14','E14','F14'].forEach((addr, idx) => {
    const c = wsDash.getCell(addr);
    c.value = g1Headers[idx];
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF475569' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });

  // Datos Tabla Gráfico 1
  const formatosData = [
    { nombre: 'Tipo B (Medianas)', ventas: 204378156.40, area: 1720242, ef: 118.81, ganador: true, color: '#10b981' },
    { nombre: 'Tipo A (Grandes)', ventas: 335957905.20, area: 3899450, ef: 86.16, ganador: false, color: '#06b6d4' },
    { nombre: 'Tipo C (Pequeñas)', ventas: 18075458.98, area: 243250, ef: 74.31, ganador: false, color: '#f59e0b' }
  ];

  formatosData.forEach((f, idx) => {
    const rNum = 15 + idx;
    const row = wsDash.getRow(rNum);
    row.height = 26;

    row.getCell('B').value = f.nombre;
    row.getCell('C').value = f.ventas;
    row.getCell('C').numFmt = '$#,##0.00';
    row.getCell('D').value = f.area;
    row.getCell('D').numFmt = '#,##0 "m²"';
    
    // Valor de Eficiencia en ROJO si es el ganador
    const efCell = row.getCell('E');
    efCell.value = f.ef;
    efCell.numFmt = '$#,##0.00';
    if (f.ganador) {
      efCell.font = redLarge;
      efCell.fill = cardBodyFill;
    } else {
      efCell.font = { name: 'Calibri', size: 11, bold: true };
    }

    // Fórmula SPARKLINE para barra visual en Google Sheets
    const sparkCell = row.getCell('F');
    sparkCell.value = {
      formula: `SPARKLINE(E${rNum}, {"charttype", "bar"; "max", 125; "color1", "${f.color}"})`,
      result: f.ef
    };

    ['B','C','D','E','F'].forEach(col => {
      row.getCell(col).border = thinBorder;
      if (!row.getCell(col).alignment) {
        row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });
  });

  // Título Gráfico 2 (Derecha)
  wsDash.mergeCells('H13:L13');
  const g2Title = wsDash.getCell('H13');
  g2Title.value = '📊 GRÁFICO 2: PARTICIPACIÓN DE VENTAS POR DEPARTAMENTO (PARETO 80/20)';
  g2Title.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  g2Title.fill = tableHeaderFill;
  g2Title.alignment = { vertical: 'middle', horizontal: 'left' };

  // Encabezados Tabla Gráfico 2
  const g2Headers = ['Departamento', 'Ventas 2012 (USD)', 'Participación (%)', 'Rol Evaluado', 'Gráfico de Participación'];
  ['H14','I14','J14','K14','L14'].forEach((addr, idx) => {
    const c = wsDash.getCell(addr);
    c.value = g2Headers[idx];
    c.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF475569' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    c.border = thinBorder;
  });

  // Datos Tabla Gráfico 2 (Ranking de Departamentos)
  const deptRanking = [
    { name: 'Despensa y Básicos', sales: 85052985.88, pct: 0.1523, tag: '⭐ DEPARTAMENTO LIDER #1', isRed: true, color: '#ef4444' },
    { name: 'Comida Fresca', sales: 59510812.89, pct: 0.1066, tag: '🟢 MOTOR DE INGRESOS', isRed: true, color: '#f97316' },
    { name: 'Artículos del Hogar y Papel', sales: 58876083.76, pct: 0.1054, tag: '🟢 MOTOR DE INGRESOS', isRed: true, color: '#f59e0b' },
    { name: 'Salud y Bienestar', sales: 51003352.42, pct: 0.0913, tag: '🟢 MOTOR DE INGRESOS', isRed: true, color: '#10b981' },
    { name: 'Ropa', sales: 40183258.61, pct: 0.0720, tag: '🟡 Consumo Estacional', isRed: false, color: '#06b6d4' },
    { name: 'Hogar y Temporada', sales: 37849580.53, pct: 0.0678, tag: '🟡 Consumo Estacional', isRed: false, color: '#06b6d4' },
    { name: 'Cuidado del Bebé y Familia', sales: 35790072.73, pct: 0.0641, tag: '🟡 Tráfico Recurrente', isRed: false, color: '#06b6d4' },
    { name: 'Snacks y Bebidas', sales: 35445927.76, pct: 0.0635, tag: '🟡 Venta Cruzada', isRed: false, color: '#06b6d4' },
    { name: 'Electrónica de Consumo', sales: 33472398.70, pct: 0.0599, tag: '🟡 Ticket Alto', isRed: false, color: '#06b6d4' },
    { name: 'Depto 16 (Sin Asignar)', sales: 29431276.26, pct: 0.0527, tag: '⚠️ Auditoría / En Revisión', isRed: false, color: '#94a3b8' },
    { name: 'Automotriz', sales: 27336093.56, pct: 0.0490, tag: '🟡 Especialidad', isRed: false, color: '#94a3b8' },
    { name: 'Cuidado de Mascotas', sales: 26532589.69, pct: 0.0475, tag: '🟡 En Crecimiento', isRed: false, color: '#94a3b8' },
    { name: 'Juguetes y Juegos', sales: 23783096.65, pct: 0.0426, tag: '⚠️ Solo Q4 (Navidad)', isRed: false, color: '#94a3b8' },
    { name: 'Oficina, Escuela y Manualidades', sales: 8187798.52, pct: 0.0147, tag: '🔴 BAJO POTENCIAL', isRed: true, color: '#dc2626' },
    { name: 'Jardín y Vida al Aire Libre', sales: 5956192.62, pct: 0.0107, tag: '🔴 BAJO POTENCIAL', isRed: true, color: '#dc2626' }
  ];

  deptRanking.forEach((d, idx) => {
    const rNum = 15 + idx;
    const row = wsDash.getRow(rNum);
    row.height = 24;

    row.getCell('H').value = d.name;
    row.getCell('I').value = d.sales;
    row.getCell('I').numFmt = '$#,##0.00';
    
    // Porcentaje de participación
    const pctCell = row.getCell('J');
    pctCell.value = d.pct;
    pctCell.numFmt = '0.00%';

    const tagCell = row.getCell('K');
    tagCell.value = d.tag;

    if (d.isRed) {
      row.getCell('H').font = redMedium;
      pctCell.font = redLarge;
      pctCell.fill = cardBodyFill;
      tagCell.font = redMedium;
    } else {
      row.getCell('H').font = { name: 'Calibri', size: 10 };
      pctCell.font = { name: 'Calibri', size: 10, bold: true };
      tagCell.font = { name: 'Calibri', size: 9, color: { argb: 'FF64748B' } };
    }

    // SPARKLINE en Google Sheets
    const sparkCell = row.getCell('L');
    sparkCell.value = {
      formula: `SPARKLINE(J${rNum}, {"charttype", "bar"; "max", 0.18; "color1", "${d.color}"})`,
      result: d.pct
    };

    ['H','I','J','K','L'].forEach(col => {
      row.getCell(col).border = thinBorder;
      if (col !== 'H') {
        row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
      } else {
        row.getCell(col).alignment = { vertical: 'middle', horizontal: 'left' };
      }
    });
  });

  // Fila Total de Participación
  const rTotDept = wsDash.getRow(30);
  rTotDept.height = 26;
  rTotDept.getCell('H').value = 'TOTAL VENTAS CADENA';
  rTotDept.getCell('H').font = { name: 'Calibri', size: 10, bold: true };
  rTotDept.getCell('I').value = { formula: 'SUM(I15:I29)', result: 558411520.58 };
  rTotDept.getCell('I').numFmt = '$#,##0.00';
  rTotDept.getCell('I').font = { name: 'Calibri', size: 10, bold: true };
  rTotDept.getCell('J').value = { formula: 'SUM(J15:J29)', result: 1 };
  rTotDept.getCell('J').numFmt = '0.00%';
  rTotDept.getCell('J').font = { name: 'Calibri', size: 10, bold: true };
  rTotDept.getCell('K').value = '100.00% Consolidado';
  rTotDept.getCell('K').font = { name: 'Calibri', size: 9, bold: true };

  ['H','I','J','K','L'].forEach(col => {
    rTotDept.getCell(col).border = boxBorder;
    rTotDept.getCell(col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
  });

  // Nota de Guía para Gráficos en Google Sheets (Fila 32)
  wsDash.mergeCells('B32:L33');
  const noteCell = wsDash.getCell('B32');
  noteCell.value = '💡 NOTA PARA LA DIRECCIÓN / EVALUACIÓN: Las barras visuales automáticas ya están generadas con SPARKLINE. Si tus profesores solicitan además insertar un gráfico flotante de Google Sheets: simplemente selecciona la tabla B14:E17 (para Gráfico 1) o H14:J29 (para Gráfico 2) y haz clic en el menú superior: Insertar > Gráfico.';
  noteCell.font = textMuted;
  noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  noteCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  noteCell.border = thinBorder;

  // Guardar archivo actualizado
  await wb.xlsx.writeFile(filePath);
  console.log('✅ Dashboard visual actualizado en:', filePath);

  const downloadPath = 'C:\\Users\\PC\\Downloads\\Proyecto 2_ Resumen Ejecutivo de Ventas Walmart  .xlsx';
  await wb.xlsx.writeFile(downloadPath);
  console.log('✅ Dashboard copiado a Descargas:', downloadPath);
}

buildDashboard().catch(console.error);
