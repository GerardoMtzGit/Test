const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const filePath = path.join(__dirname, '..', 'data', 'Proyecto 2_ Resumen Ejecutivo de Ventas Walmart.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  console.log('Hojas en el workbook:', wb.worksheets.map(w => w.name));

  // Estilos de Letras Rojas Grandes para Respuestas Evaluadas
  const redLargeFont = {
    name: 'Calibri',
    size: 14,
    bold: true,
    color: { argb: 'FFFF0000' } // Rojo brillante
  };

  const redHeroFont = {
    name: 'Calibri',
    size: 16,
    bold: true,
    color: { argb: 'FFCC0000' } // Rojo oscuro destacado
  };

  const redMediumFont = {
    name: 'Calibri',
    size: 12,
    bold: true,
    color: { argb: 'FFFF0000' }
  };

  const headerFont = {
    name: 'Calibri',
    size: 11,
    bold: true,
    color: { argb: 'FFFFFFFF' }
  };

  const headerFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' } // Azul oscuro elegante
  };

  const yellowHighlightFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFBEB' } // Fondo suave para destacar
  };

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };

  // ==========================================
  // 1. HOJA RESUMEN (C-F-I)
  // ==========================================
  const wsResumen = wb.getWorksheet('Resumen');
  if (wsResumen) {
    wsResumen.columns = [
      { width: 32 }, // Pregunta
      { width: 28 }, // KPI
      { width: 35 }, // Contexto
      { width: 45 }, // Insight
      { width: 45 }  // Implicación
    ];

    // Estilo encabezado
    const rowHeader = wsResumen.getRow(1);
    rowHeader.height = 28;
    rowHeader.eachCell(cell => {
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = thinBorder;
    });

    // Fila 2 (Pregunta 1: Eficiencia)
    const r2 = wsResumen.getRow(2);
    r2.height = 110;
    r2.getCell(1).alignment = { vertical: 'middle', wrapText: true };
    r2.getCell(2).alignment = { vertical: 'middle', wrapText: true };
    r2.getCell(3).alignment = { vertical: 'middle', wrapText: true };
    
    // RESPUESTAS EVALUADAS EN ROJO GRANDE
    r2.getCell(4).font = redLargeFont;
    r2.getCell(4).fill = yellowHighlightFill;
    r2.getCell(4).alignment = { vertical: 'middle', wrapText: true };
    r2.getCell(4).border = thinBorder;

    r2.getCell(5).font = redLargeFont;
    r2.getCell(5).fill = yellowHighlightFill;
    r2.getCell(5).alignment = { vertical: 'middle', wrapText: true };
    r2.getCell(5).border = thinBorder;

    // Fila 3 (Pregunta 2: Participación)
    const r3 = wsResumen.getRow(3);
    r3.height = 110;
    r3.getCell(1).alignment = { vertical: 'middle', wrapText: true };
    r3.getCell(2).alignment = { vertical: 'middle', wrapText: true };
    r3.getCell(3).alignment = { vertical: 'middle', wrapText: true };

    // RESPUESTAS EVALUADAS EN ROJO GRANDE
    r3.getCell(4).font = redLargeFont;
    r3.getCell(4).fill = yellowHighlightFill;
    r3.getCell(4).alignment = { vertical: 'middle', wrapText: true };
    r3.getCell(4).border = thinBorder;

    r3.getCell(5).font = redLargeFont;
    r3.getCell(5).fill = yellowHighlightFill;
    r3.getCell(5).alignment = { vertical: 'middle', wrapText: true };
    r3.getCell(5).border = thinBorder;
  }

  // ==========================================
  // 2. HOJA DASHBOARD
  // ==========================================
  const wsDash = wb.getWorksheet('Dashboard');
  if (wsDash) {
    wsDash.columns = [
      { width: 34 }, // Indicador
      { width: 28 }, // Valor
      { width: 42 }, // Fórmula en español
      { width: 50 }  // Interpretación
    ];

    // Título Dashboard
    wsDash.getRow(1).height = 32;
    wsDash.getCell('A1').font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FF1E293B' } };

    // Fila de encabezado tabla (Fila 4)
    const dashHeader = wsDash.getRow(4);
    dashHeader.height = 26;
    dashHeader.eachCell(cell => {
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = thinBorder;
    });

    // Actualizar fórmulas en español explícitas en la columna C
    wsDash.getCell('C5').value = '=SUMA(clean_ventas!H:H)';
    wsDash.getCell('C6').value = '=CONTARA(raw_tiendas!A2:A46)';
    wsDash.getCell('C7').value = '=SUMAR.SI(clean_ventas!B:B, "B", clean_ventas!H:H) / SUMAR.SI(raw_tiendas!B:B, "B", raw_tiendas!C:C)';
    wsDash.getCell('C8').value = '=INDICE(Pivot!A10:A24, COINCIDIR(MAX(Pivot!B10:B24), Pivot!B10:B24, 0))';
    wsDash.getCell('C9').value = '=SUMA(Pivot!C10:C13)';
    wsDash.getCell('C10').value = '=CONTAR(clean_ventas!H:H)';

    // ESTILIZAR RESPUESTAS EVALUADAS EN ROJO GRANDE
    const evaluatedRows = [5, 6, 7, 8, 9, 10];
    evaluatedRows.forEach(rowNum => {
      const row = wsDash.getRow(rowNum);
      row.height = 36;
      row.eachCell(cell => {
        cell.border = thinBorder;
      });

      // El valor clave en ROJO GRANDE
      const valCell = row.getCell(2);
      valCell.font = redHeroFont;
      valCell.fill = yellowHighlightFill;
      valCell.alignment = { vertical: 'middle', horizontal: 'center' };

      // La fórmula en español destacada
      const formulaCell = row.getCell(3);
      formulaCell.font = { name: 'Consolas', size: 11, bold: true, color: { argb: 'FF0284C7' } };
      formulaCell.alignment = { vertical: 'middle', horizontal: 'left' };
    });
  }

  // ==========================================
  // 3. HOJA PIVOT
  // ==========================================
  const wsPivot = wb.getWorksheet('Pivot');
  if (wsPivot) {
    wsPivot.columns = [
      { width: 34 },
      { width: 28 },
      { width: 24 },
      { width: 24 },
      { width: 36 } // Columna extra para etiqueta de evaluación
    ];

    // Título Reporte 1
    wsPivot.getCell('A1').font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FF0F172A' } };

    // Encabezado Reporte 1 (Fila 2)
    const h1 = wsPivot.getRow(2);
    h1.height = 24;
    h1.eachCell(c => { c.font = headerFont; c.fill = headerFill; c.alignment = { vertical: 'middle', horizontal: 'center' }; c.border = thinBorder; });
    wsPivot.getCell('E2').value = 'EVALUACION / HALLAZGO';
    wsPivot.getCell('E2').font = headerFont;
    wsPivot.getCell('E2').fill = headerFill;

    // Fila 4 (Tipo B - GANADOR) EN ROJO GRANDE
    const rTipoB = wsPivot.getRow(4);
    rTipoB.height = 32;
    rTipoB.getCell(1).font = redLargeFont;
    rTipoB.getCell(2).font = redLargeFont;
    rTipoB.getCell(3).font = redLargeFont;
    rTipoB.getCell(4).font = redHeroFont; // $118.81 en rojo grande
    rTipoB.getCell(4).fill = yellowHighlightFill;
    rTipoB.getCell(5).value = '🏆 FORMATO MAS EFICIENTE (+37.9% vs Tipo A)';
    rTipoB.getCell(5).font = redLargeFont;
    rTipoB.getCell(5).fill = yellowHighlightFill;

    // Título Reporte 2
    wsPivot.getCell('A8').font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FF0F172A' } };

    // Encabezado Reporte 2 (Fila 9)
    const h2 = wsPivot.getRow(9);
    h2.height = 24;
    h2.eachCell(c => { c.font = headerFont; c.fill = headerFill; c.alignment = { vertical: 'middle', horizontal: 'center' }; c.border = thinBorder; });
    wsPivot.getCell('D9').value = 'ROL EN EL NEGOCIO (EVALUACION)';
    wsPivot.getCell('D9').font = headerFont;
    wsPivot.getCell('D9').fill = headerFill;

    // Despensa y Básicos (#1) EN ROJO GRANDE
    const rDespensa = wsPivot.getRow(10);
    rDespensa.height = 30;
    rDespensa.getCell(1).font = redLargeFont;
    rDespensa.getCell(2).font = redLargeFont;
    rDespensa.getCell(3).font = redHeroFont; // 15.23% en rojo grande
    rDespensa.getCell(3).fill = yellowHighlightFill;
    rDespensa.getCell(4).value = '⭐ DEPARTAMENTO LIDER #1 (15.23% del Total)';
    rDespensa.getCell(4).font = redLargeFont;
    rDespensa.getCell(4).fill = yellowHighlightFill;

    // Top 2, 3, 4 en Rojo Mediano
    [11, 12, 13].forEach((rNum, idx) => {
      const r = wsPivot.getRow(rNum);
      r.height = 24;
      r.getCell(1).font = redMediumFont;
      r.getCell(2).font = redMediumFont;
      r.getCell(3).font = redMediumFont;
      r.getCell(4).value = `🟢 MOTOR DE INGRESOS (Top ${idx + 2})`;
      r.getCell(4).font = redMediumFont;
    });

    // Rezagados (Jardín y Oficina) en Rojo
    const rOficina = wsPivot.getRow(23); // Oficina
    rOficina.getCell(1).font = redMediumFont;
    rOficina.getCell(3).font = redMediumFont;
    rOficina.getCell(4).value = '⚠️ BAJO POTENCIAL (1.47%)';
    rOficina.getCell(4).font = redMediumFont;

    const rJardin = wsPivot.getRow(24); // Jardín
    rJardin.getCell(1).font = redMediumFont;
    rJardin.getCell(3).font = redMediumFont;
    rJardin.getCell(4).value = '⚠️ MENOR APORTACION (1.07%)';
    rJardin.getCell(4).font = redMediumFont;
  }

  // ==========================================
  // 4. HOJA README (QA CHECKS)
  // ==========================================
  const wsReadme = wb.getWorksheet('README');
  if (wsReadme) {
    wsReadme.columns = [
      { width: 34 },
      { width: 38 },
      { width: 44 },
      { width: 48 }
    ];

    // Buscar la fila de QA / Validaciones y colorear los resultados en ROJO GRANDE
    wsReadme.eachRow((row, rowNumber) => {
      const valA = String(row.getCell(1).value || '');
      if (valA.includes('¿Existen departamentos') || valA.includes('¿Existen ventas') || valA.includes('¿Metros cuadrados')) {
        row.height = 34;
        const resCell = row.getCell(4);
        resCell.font = redLargeFont;
        resCell.fill = yellowHighlightFill;
        resCell.alignment = { vertical: 'middle', wrapText: true };

        const formCell = row.getCell(3);
        formCell.font = { name: 'Consolas', size: 11, bold: true, color: { argb: 'FF0284C7' } };
      }
    });
  }

  // Guardar archivo actualizado en data y downloads
  await wb.xlsx.writeFile(filePath);
  console.log('✅ Archivo guardado con letras rojas grandes en:', filePath);

  const downloadPath = 'C:\\Users\\PC\\Downloads\\Proyecto 2_ Resumen Ejecutivo de Ventas Walmart  .xlsx';
  await wb.xlsx.writeFile(downloadPath);
  console.log('✅ Archivo copiado con éxito a Descargas:', downloadPath);
}

main().catch(console.error);
