// ============================================================
//  EXCEL MANAGER - Guarda y organiza gastos en .xlsx
//  Una hoja por proyecto, resumen general en "Dashboard"
// ============================================================

const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const EXCEL_PATH = process.env.EXCEL_PATH || path.join(__dirname, "../gastos_empresa.xlsx");

// Colores corporativos
const COLORS = {
  header: "1F3864",       // Azul oscuro
  headerText: "FFFFFF",
  accent: "2E75B6",       // Azul
  evenRow: "EBF3FB",
  oddRow: "FFFFFF",
  total: "C6EFCE",        // Verde claro
  border: "BDD7EE",
  dashboardHeader: "203864",
};

const CATEGORIES_COLORS = {
  Comida: "FF6B6B",
  Transporte: "4ECDC4",
  Hospedaje: "45B7D1",
  Material: "96CEB4",
  Servicios: "FFEAA7",
  Software: "DDA0DD",
  Marketing: "F0A500",
  Otro: "B0BEC5",
};

// ─── Función principal: guarda un gasto ─────────────────────
async function saveExpenseToExcel(expenseData) {
  let workbook = new ExcelJS.Workbook();

  // Cargar archivo existente o crear nuevo
  if (fs.existsSync(EXCEL_PATH)) {
    await workbook.xlsx.readFile(EXCEL_PATH);
  } else {
    console.log("📄 Creando nuevo archivo Excel...");
  }

  // 1. Guardar en hoja del proyecto específico
  await saveToProjectSheet(workbook, expenseData);

  // 2. Guardar en hoja "Todos los Gastos"
  await saveToAllExpensesSheet(workbook, expenseData);

  // 3. Actualizar Dashboard
  await updateDashboard(workbook);

  // Guardar archivo
  await workbook.xlsx.writeFile(EXCEL_PATH);
  console.log(`💾 Excel guardado: ${EXCEL_PATH}`);

  return EXCEL_PATH;
}

// ─── Hoja por proyecto ───────────────────────────────────────
async function saveToProjectSheet(workbook, data) {
  const sheetName = sanitizeSheetName(data.project);
  let sheet = workbook.getWorksheet(sheetName);

  if (!sheet) {
    sheet = workbook.addWorksheet(sheetName, {
      properties: { tabColor: { argb: COLORS.accent } },
    });
    createProjectSheetHeaders(sheet);
  }

  // Agregar fila de datos
  const rowIndex = sheet.rowCount + 1;
  const row = sheet.addRow([
    data.id,
    data.date || new Date().toISOString().split("T")[0],
    data.vendor || "N/A",
    data.category || "Otro",
    data.subtotal || 0,
    data.tax || 0,
    data.total || 0,
    data.currency || "MXN",
    data.paymentMethod || "N/A",
    (data.items || []).join(", "),
    data.sentBy,
    data.notes || "",
  ]);

  // Estilo de fila alterna
  styleDataRow(row, rowIndex, data.category);

  // Actualizar fórmula de totales
  updateProjectTotals(sheet);
}

// ─── Headers de hoja de proyecto ────────────────────────────
function createProjectSheetHeaders(sheet) {
  // Título del proyecto
  sheet.mergeCells("A1:L1");
  const titleCell = sheet.getCell("A1");
  titleCell.value = `📊 GASTOS - ${sheet.name.toUpperCase()}`;
  titleCell.font = { name: "Calibri", size: 16, bold: true, color: { argb: COLORS.headerText } };
  titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.dashboardHeader } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 35;

  // Fila de headers
  const headers = [
    { header: "ID", key: "id", width: 18 },
    { header: "Fecha", key: "date", width: 14 },
    { header: "Proveedor", key: "vendor", width: 25 },
    { header: "Categoría", key: "category", width: 16 },
    { header: "Subtotal", key: "subtotal", width: 14 },
    { header: "IVA", key: "tax", width: 12 },
    { header: "Total", key: "total", width: 14 },
    { header: "Moneda", key: "currency", width: 10 },
    { header: "Método Pago", key: "payment", width: 16 },
    { header: "Concepto", key: "items", width: 35 },
    { header: "Registrado por", key: "sentBy", width: 18 },
    { header: "Notas", key: "notes", width: 25 },
  ];

  sheet.columns = headers;
  const headerRow = sheet.getRow(2);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h.header;
    cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.headerText } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.header } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: COLORS.border } },
      bottom: { style: "thin", color: { argb: COLORS.border } },
      left: { style: "thin", color: { argb: COLORS.border } },
      right: { style: "thin", color: { argb: COLORS.border } },
    };
  });
  headerRow.height = 28;

  // Auto-filter
  sheet.autoFilter = { from: "A2", to: "L2" };
  // Fijar header
  sheet.views = [{ state: "frozen", ySplit: 2 }];
}

// ─── Estilo de fila de datos ─────────────────────────────────
function styleDataRow(row, rowIndex, category) {
  const isEven = rowIndex % 2 === 0;
  const bgColor = isEven ? COLORS.evenRow : COLORS.oddRow;

  row.eachCell((cell, colNumber) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgColor } };
    cell.border = {
      top: { style: "hair", color: { argb: COLORS.border } },
      bottom: { style: "hair", color: { argb: COLORS.border } },
      left: { style: "hair", color: { argb: COLORS.border } },
      right: { style: "hair", color: { argb: COLORS.border } },
    };
    cell.font = { name: "Calibri", size: 10 };
    cell.alignment = { vertical: "middle", wrapText: true };

    // Formato de moneda para columnas numéricas (5, 6, 7)
    if ([5, 6, 7].includes(colNumber)) {
      cell.numFmt = '#,##0.00';
      cell.alignment = { horizontal: "right", vertical: "middle" };
    }

    // Color de categoría en columna 4
    if (colNumber === 4 && category && CATEGORIES_COLORS[category]) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: CATEGORIES_COLORS[category].replace("#", "") },
      };
      cell.font = { name: "Calibri", size: 10, bold: true };
    }
  });
  row.height = 22;
}

// ─── Totales al final de hoja ────────────────────────────────
function updateProjectTotals(sheet) {
  const lastDataRow = sheet.rowCount;
  if (lastDataRow < 3) return;

  // Eliminar fila de totales anterior si existe
  const possibleTotalsRow = sheet.getRow(lastDataRow);
  if (possibleTotalsRow.getCell(1).value === "TOTAL") {
    sheet.spliceRows(lastDataRow, 1);
  }

  const totalsRow = sheet.addRow([
    "TOTAL", "", "", "",
    { formula: `SUM(E3:E${sheet.rowCount})` },
    { formula: `SUM(F3:F${sheet.rowCount})` },
    { formula: `SUM(G3:G${sheet.rowCount})` },
    "", "", "", "", "",
  ]);

  totalsRow.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.total } };
    cell.font = { name: "Calibri", size: 11, bold: true };
    cell.border = {
      top: { style: "medium", color: { argb: "217346" } },
      bottom: { style: "medium", color: { argb: "217346" } },
    };
  });
  totalsRow.getCell(1).alignment = { horizontal: "center" };
}

// ─── Hoja "Todos los Gastos" ─────────────────────────────────
async function saveToAllExpensesSheet(workbook, data) {
  let sheet = workbook.getWorksheet("Todos los Gastos");
  if (!sheet) {
    sheet = workbook.addWorksheet("Todos los Gastos", {
      properties: { tabColor: { argb: "2E75B6" } },
    });

    // Headers para hoja general
    sheet.columns = [
      { header: "ID", key: "id", width: 18 },
      { header: "Fecha", key: "date", width: 14 },
      { header: "Proyecto", key: "project", width: 20 },
      { header: "Proveedor", key: "vendor", width: 25 },
      { header: "Categoría", key: "category", width: 16 },
      { header: "Total", key: "total", width: 14 },
      { header: "Moneda", key: "currency", width: 10 },
      { header: "Método Pago", key: "payment", width: 16 },
      { header: "Registrado por", key: "sentBy", width: 18 },
    ];

    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.headerText } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.header } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });
    headerRow.height = 28;
    sheet.autoFilter = { from: "A1", to: "I1" };
    sheet.views = [{ state: "frozen", ySplit: 1 }];
  }

  const rowIndex = sheet.rowCount + 1;
  const row = sheet.addRow([
    data.id, data.date, data.project, data.vendor,
    data.category, data.total, data.currency,
    data.paymentMethod, data.sentBy,
  ]);
  styleDataRow(row, rowIndex, data.category);
}

// ─── Dashboard con resumen ───────────────────────────────────
async function updateDashboard(workbook) {
  let dash = workbook.getWorksheet("Dashboard");
  if (!dash) {
    dash = workbook.addWorksheet("Dashboard", {
      properties: { tabColor: { argb: "1F3864" } },
    });
  }

  dash.getColumn(1).width = 30;
  dash.getColumn(2).width = 20;

  // Título
  dash.mergeCells("A1:B1");
  const title = dash.getCell("A1");
  title.value = "📊 DASHBOARD DE GASTOS";
  title.font = { name: "Calibri", size: 20, bold: true, color: { argb: COLORS.headerText } };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.dashboardHeader } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  dash.getRow(1).height = 45;

  // Última actualización
  dash.getCell("A2").value = `Última actualización: ${new Date().toLocaleString("es-MX")}`;
  dash.getCell("A2").font = { name: "Calibri", size: 10, italic: true, color: { argb: "666666" } };

  // Resumen por proyecto
  dash.getCell("A4").value = "PROYECTO";
  dash.getCell("B4").value = "TOTAL GASTADO (MXN)";
  ["A4", "B4"].forEach((cell) => {
    dash.getCell(cell).font = { name: "Calibri", bold: true, color: { argb: COLORS.headerText } };
    dash.getCell(cell).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.accent } };
    dash.getCell(cell).alignment = { horizontal: "center" };
  });

  // Listar proyectos existentes
  let row = 5;
  workbook.eachSheet((sheet) => {
    if (!["Dashboard", "Todos los Gastos"].includes(sheet.name)) {
      dash.getCell(`A${row}`).value = sheet.name;
      dash.getCell(`A${row}`).font = { name: "Calibri", size: 11 };
      // Referencia a suma de la columna G del proyecto
      dash.getCell(`B${row}`).value = {
        formula: `SUMIF('Todos los Gastos'!C:C,"${sheet.name}",'Todos los Gastos'!F:F)`,
      };
      dash.getCell(`B${row}`).numFmt = '$#,##0.00';
      dash.getCell(`B${row}`).font = { name: "Calibri", size: 11, bold: true, color: { argb: "217346" } };
      row++;
    }
  });

  // Total general
  dash.getCell(`A${row}`).value = "TOTAL GENERAL";
  dash.getCell(`B${row}`).value = { formula: `SUM(B5:B${row - 1})` };
  dash.getCell(`B${row}`).numFmt = '$#,##0.00';
  [`A${row}`, `B${row}`].forEach((cell) => {
    dash.getCell(cell).font = { name: "Calibri", size: 13, bold: true, color: { argb: COLORS.headerText } };
    dash.getCell(cell).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.header } };
  });
}

// ─── Utilidades ──────────────────────────────────────────────
function sanitizeSheetName(name) {
  return name
    .replace(/[\\/*?:[\]]/g, "")
    .substring(0, 31)
    .trim() || "Sin Proyecto";
}

module.exports = { saveExpenseToExcel };
