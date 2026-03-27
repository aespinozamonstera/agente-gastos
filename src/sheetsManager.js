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
        requests: [{ addSheet: { properties: { title: sheetName, index: 0 } } }],
      },
    });
  }

  // Leer todos los gastos
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Todos los Gastos!A2:I1000",
  });

  const rows = response.data.values || [];

  // Calcular totales por proyecto
  const porProyecto = {};
  const porCategoria = {};
  const porMes = {};
  let totalGasolina = 0;

  rows.forEach(row => {
    const fecha = row[1] || "";
    const proyecto = row[2] || "Sin Proyecto";
    const categoria = row[4] || "Otro";
    const total = parseFloat(row[5]) || 0;
    const mes = fecha.substring(0, 7); // YYYY-MM

    porProyecto[proyecto] = (porProyecto[proyecto] || 0) + total;
    porCategoria[categoria] = (porCategoria[categoria] || 0) + total;
    if (mes) porMes[mes] = (porMes[mes] || 0) + total;
    if (categoria === "Transporte") totalGasolina += total;
  });

  // Construir datos del dashboard
  const dashData = [
    ["📊 DASHBOARD DE GASTOS - MONSTERA ESTUDIO"],
    [`Actualizado: ${new Date().toLocaleString("es-MX")}`],
    [""],
    ["📁 TOTAL POR PROYECTO", "", "🏷️ POR CATEGORÍA", "", "📅 POR MES"],
    ["Proyecto", "Total MXN", "Categoría", "Total MXN", "Mes", "Total MXN"],
  ];

  const proyectos = Object.entries(porProyecto).sort((a, b) => b[1] - a[1]);
  const categorias = Object.entries(porCategoria).sort((a, b) => b[1] - a[1]);
  const meses = Object.entries(porMes).sort((a, b) => a[0].localeCompare(b[0]));

  const maxRows = Math.max(proyectos.length, categorias.length, meses.length);

  for (let i = 0; i < maxRows; i++) {
    dashData.push([
      proyectos[i] ? proyectos[i][0] : "",
      proyectos[i] ? proyectos[i][1].toFixed(2) : "",
      categorias[i] ? categorias[i][0] : "",
      categorias[i] ? categorias[i][1].toFixed(2) : "",
      meses[i] ? meses[i][0] : "",
      meses[i] ? meses[i][1].toFixed(2) : "",
    ]);
  }

  dashData.push([""]);
  dashData.push(["⛽ TOTAL GASOLINA/TRANSPORTE", totalGasolina.toFixed(2)]);

  // Limpiar y escribir
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1:Z100`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: dashData },
  });

  console.log("📊 Dashboard actualizado");
}
module.exports = { saveExpenseToExcel };
