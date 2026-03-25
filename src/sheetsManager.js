// ============================================================
//  GOOGLE SHEETS MANAGER - Guarda gastos en Google Sheets
// ============================================================

const { google } = require("googleapis");

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Categorías y sus colores
const CATEGORY_COLORS = {
  Comida:      { red: 1, green: 0.42, blue: 0.42 },
  Transporte:  { red: 0.31, green: 0.80, blue: 0.77 },
  Hospedaje:   { red: 0.27, green: 0.72, blue: 0.82 },
  Material:    { red: 0.59, green: 0.81, blue: 0.71 },
  Servicios:   { red: 1, green: 0.92, blue: 0.65 },
  Software:    { red: 0.87, green: 0.63, blue: 0.87 },
  Marketing:   { red: 0.94, green: 0.65, blue: 0 },
  Otro:        { red: 0.69, green: 0.75, blue: 0.77 },
};

// ─── Autenticación con Google ────────────────────────────────
function getAuth() {
  const b64 = process.env.GOOGLE_CREDENTIALS_B64;
  console.log('B64 primeros 20:', b64 ? b64.substring(0, 20) : 'VACIO');
  const decoded = Buffer.from(b64, 'base64').toString('utf8');
  console.log('JSON primeros 20:', decoded.substring(0, 20));
  const credentials = JSON.parse(decoded);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
  });
}

// ─── Función principal: guarda un gasto ─────────────────────
async function saveExpenseToExcel(expenseData) {
  const auth = getAuth();
  const sheets = google.sheets({ version: "v4", auth });

  // 1. Guardar en hoja del proyecto
  await saveToProjectSheet(sheets, expenseData);

  // 2. Guardar en hoja "Todos los Gastos"
  await saveToAllExpensesSheet(sheets, expenseData);
await updateDashboard(sheets);
  console.log(`💾 Guardado en Google Sheets: ${expenseData.vendor}`);
}

// ─── Guardar en hoja del proyecto ───────────────────────────
async function saveToProjectSheet(sheets, data) {
  const sheetName = data.project.substring(0, 31);

  // Verificar si la hoja existe
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const existingSheets = spreadsheet.data.sheets.map(
    (s) => s.properties.title
  );

  // Crear hoja si no existe
  if (!existingSheets.includes(sheetName)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName,
                tabColor: { red: 0.18, green: 0.46, blue: 0.71 },
              },
            },
          },
        ],
      },
    });

    // Agregar headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1:L1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "ID", "Fecha", "Proveedor", "Categoría",
          "Subtotal", "IVA", "Total", "Moneda",
          "Método Pago", "Concepto", "Registrado por", "Notas"
        ]],
      },
    });

    // Formato de headers
    const sheetId = await getSheetId(sheets, sheetName);
    await formatHeaders(sheets, sheetId);
  }

  // Agregar fila de datos
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:L`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
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
      ]],
    },
  });
}

// ─── Guardar en hoja "Todos los Gastos" ─────────────────────
async function saveToAllExpensesSheet(sheets, data) {
  const sheetName = "Todos los Gastos";

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });
  const existingSheets = spreadsheet.data.sheets.map(
    (s) => s.properties.title
  );

  if (!existingSheets.includes(sheetName)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: sheetName,
              index: 0,
              tabColor: { red: 0.13, green: 0.29, blue: 0.53 },
            },
          },
        }],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1:I1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "ID", "Fecha", "Proyecto", "Proveedor",
          "Categoría", "Total", "Moneda",
          "Método Pago", "Registrado por"
        ]],
      },
    });

    const sheetId = await getSheetId(sheets, sheetName);
    await formatHeaders(sheets, sheetId);
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:I`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        data.id,
        data.date,
        data.project,
        data.vendor,
        data.category,
        data.total,
        data.currency,
        data.paymentMethod,
        data.sentBy,
      ]],
    },
  });
}

// ─── Obtener ID de una hoja por nombre ──────────────────────
async function getSheetId(sheets, sheetName) {
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });
  const sheet = spreadsheet.data.sheets.find(
    (s) => s.properties.title === sheetName
  );
  return sheet?.properties?.sheetId;
}

// ─── Formato de headers ──────────────────────────────────────
async function formatHeaders(sheets, sheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.12, green: 0.22, blue: 0.39 },
                textFormat: {
                  foregroundColor: { red: 1, green: 1, blue: 1 },
                  bold: true,
                  fontSize: 11,
                },
                horizontalAlignment: "CENTER",
              },
            },
            fields: "userEnteredFormat",
          },
        },
        {
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "gridProperties.frozenRowCount",
          },
        },
      ],
    },
  });
}
// ─── Actualizar Dashboard ────────────────────────────────────
async function updateDashboard(sheets) {
  const sheetName = "Dashboard";
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });
  const existingSheets = spreadsheet.data.sheets.map(s => s.properties.title);

  if (!existingSheets.includes(sheetName)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: sheetName,
              index: 0,
              tabColor: { red: 0.13, green: 0.29, blue: 0.53 },
            },
          },
        }],
      },
    });
  }

  // Limpiar hoja
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1:Z100`,
  });

  // Título
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [["📊 DASHBOARD DE GASTOS - MONSTERA ESTUDIO"]] },
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A2`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[`Actualizado: ${new Date().toLocaleString("es-MX")}`]] },
  });

  // Sección 1: Total por Proyecto
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A4`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["📁 TOTAL POR PROYECTO"],
        ["Proyecto", "Total (MXN)", "# Gastos"],
        [
          { formula: `=IFERROR(UNIQUE(FILTER('Todos los Gastos'!C2:C, 'Todos los Gastos'!C2:C<>"")), "Sin datos")` },
          "",
          "",
        ],
      ],
    },
  });

  // Fórmulas por proyecto
  for (let i = 0; i < 10; i++) {
    const row = 7 + i;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!B${row}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          { formula: `=IFERROR(SUMIF('Todos los Gastos'!C:C,A${row},'Todos los Gastos'!F:F),0)` },
          { formula: `=IFERROR(COUNTIF('Todos los Gastos'!C:C,A${row}),0)` },
        ]],
      },
    });
  }

  // Sección 2: Por Categoría
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!E4`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["🏷️ GASTOS POR CATEGORÍA"],
        ["Categoría", "Total (MXN)", "# Gastos"],
        ["Comida", { formula: `=SUMIF('Todos los Gastos'!E:E,"Comida",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Comida")` }],
        ["Transporte", { formula: `=SUMIF('Todos los Gastos'!E:E,"Transporte",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Transporte")` }],
        ["Hospedaje", { formula: `=SUMIF('Todos los Gastos'!E:E,"Hospedaje",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Hospedaje")` }],
        ["Material", { formula: `=SUMIF('Todos los Gastos'!E:E,"Material",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Material")` }],
        ["Servicios", { formula: `=SUMIF('Todos los Gastos'!E:E,"Servicios",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Servicios")` }],
        ["Software", { formula: `=SUMIF('Todos los Gastos'!E:E,"Software",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Software")` }],
        ["Marketing", { formula: `=SUMIF('Todos los Gastos'!E:E,"Marketing",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Marketing")` }],
        ["Otro", { formula: `=SUMIF('Todos los Gastos'!E:E,"Otro",'Todos los Gastos'!F:F)` }, { formula: `=COUNTIF('Todos los Gastos'!E:E,"Otro")` }],
        ["TOTAL", { formula: `=SUM(F6:F13)` }, { formula: `=SUM(G6:G13)` }],
      ],
    },
  });

  // Sección 3: Por Mes
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A20`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["📅 GASTOS POR MES"],
        ["Mes", "Total (MXN)", "# Gastos"],
        ["Enero", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=1))` }, ""],
        ["Febrero", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=2))` }, ""],
        ["Marzo", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=3))` }, ""],
        ["Abril", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=4))` }, ""],
        ["Mayo", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=5))` }, ""],
        ["Junio", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=6))` }, ""],
        ["Julio", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=7))` }, ""],
        ["Agosto", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=8))` }, ""],
        ["Septiembre", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=9))` }, ""],
        ["Octubre", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=10))` }, ""],
        ["Noviembre", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=11))` }, ""],
        ["Diciembre", { formula: `=SUMPRODUCT(('Todos los Gastos'!F2:F1000)*( MONTH(DATEVALUE('Todos los Gastos'!B2:B1000))=12))` }, ""],
      ],
    },
  });

  // Sección 4: Gasolinas
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!E20`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["⛽ GASTOS DE GASOLINA"],
        ["Proyecto", "Total Gasolina"],
        ["TOTAL", { formula: `=SUMIF('Todos los Gastos'!E:E,"Transporte",'Todos los Gastos'!F:F)` }],
      ],
    },
  });

  console.log("📊 Dashboard actualizado");
}
module.exports = { saveExpenseToExcel };
