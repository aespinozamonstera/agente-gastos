// ============================================================
//  AGENTE DE GASTOS - Servidor con Telegram
//  Recibe fotos de tickets por Telegram → Claude Vision → Excel
// ============================================================

require("dotenv").config();
const express = require("express");
const axios = require("axios");
const Anthropic = require("@anthropic-ai/sdk");
const { saveExpenseToExcel } = require("./excelManager");

const app = express();
app.use(express.json());

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const TELEGRAM_TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const TELEGRAM_API = `https://api.telegram.org/bot${TELEGRAM_TOKEN}`;

// ─── Webhook: Telegram manda aquí cada mensaje ───────────────
app.post("/webhook", async (req, res) => {
  res.sendStatus(200); // Siempre responder rápido a Telegram

  try {
    const update = req.body;
    const message = update.message;
    if (!message) return;

    const chatId = message.chat.id;
    const firstName = message.chat.first_name || "amigo";

    // ── El usuario mandó una FOTO ────────────────────────────
    if (message.photo) {
      const caption = message.caption || "";

      // Detectar proyecto desde el pie de foto
      const projectMatch = caption.match(/proyecto[:\s]+([^\n,]+)/i);
      const project = projectMatch ? projectMatch[1].trim() : "Sin Proyecto";

      await sendMessage(chatId, `⏳ Procesando tu ticket para el proyecto *${project}*...`);

      // Obtener la foto en mejor calidad (última = mayor resolución)
      const bestPhoto = message.photo[message.photo.length - 1];
      const imageBase64 = await downloadTelegramImage(bestPhoto.file_id);

      // Analizar con Claude Vision
      const expenseData = await analyzeReceiptWithClaude(imageBase64, project, firstName);

      // Guardar en Excel
      await saveExpenseToExcel(expenseData);

      // Confirmar al usuario
      await sendMessage(chatId, formatConfirmationMessage(expenseData));

      console.log(`✅ Gasto guardado: $${expenseData.total} - ${expenseData.vendor}`);
    }

    // ── Comando /start ────────────────────────────────────────
    else if (message.text && message.text.startsWith("/start")) {
      await sendMessage(chatId,
        `👋 ¡Hola ${firstName}! Soy tu *Agente de Gastos*.\n\n` +
        `📸 Mándame una foto de cualquier ticket o recibo y lo registro automáticamente en Excel.\n\n` +
        `📁 Para indicar el proyecto, escribe en el pie de foto:\n` +
        `_Proyecto: Nombre del Proyecto_\n\n` +
        `Si no indicas proyecto, se guardará en "Sin Proyecto".`
      );
    }

    // ── Cualquier otro texto ──────────────────────────────────
    else if (message.text) {
      await sendMessage(chatId,
        `📸 Para registrar un gasto, *envíame una foto del ticket*.\n\n` +
        `Agrega el proyecto en el pie de foto:\n` +
        `_Ej: Proyecto: ACME Corp_`
      );
    }

  } catch (error) {
    console.error("❌ Error procesando mensaje:", error.message);
  }
});

// ─── Descargar imagen desde Telegram ────────────────────────
async function downloadTelegramImage(fileId) {
  const fileResponse = await axios.get(`${TELEGRAM_API}/getFile?file_id=${fileId}`);
  const filePath = fileResponse.data.result.file_path;
  const imageResponse = await axios.get(
    `https://api.telegram.org/file/bot${TELEGRAM_TOKEN}/${filePath}`,
    { responseType: "arraybuffer" }
  );
  return Buffer.from(imageResponse.data).toString("base64");
}

// ─── Analizar ticket con Claude Vision ──────────────────────
async function analyzeReceiptWithClaude(imageBase64, project, sentBy) {
  const response = await anthropic.messages.create({
    model: "claude-opus-4-5",
    max_tokens: 1024,
    messages: [{
      role: "user",
      content: [
        {
          type: "image",
          source: { type: "base64", media_type: "image/jpeg", data: imageBase64 },
        },
        {
          type: "text",
          text: `Analiza este ticket/recibo y responde ÚNICAMENTE con JSON (sin texto extra):
{
  "vendor": "nombre del negocio",
  "date": "YYYY-MM-DD",
  "total": 0.00,
  "currency": "MXN",
  "category": "Comida|Transporte|Hospedaje|Material|Servicios|Software|Marketing|Otro",
  "items": ["descripción breve"],
  "tax": 0.00,
  "subtotal": 0.00,
  "paymentMethod": "efectivo/tarjeta/transferencia o null",
  "notes": null
}`,
        },
      ],
    }],
  });

  const parsed = JSON.parse(response.content[0].text.trim());
  return {
    ...parsed,
    project,
    sentBy,
    registeredAt: new Date().toISOString(),
    id: `EXP-${Date.now()}`,
  };
}

// ─── Enviar mensaje a Telegram ───────────────────────────────
async function sendMessage(chatId, text) {
  await axios.post(`${TELEGRAM_API}/sendMessage`, {
    chat_id: chatId,
    text,
    parse_mode: "Markdown",
  });
}

// ─── Mensaje de confirmación ─────────────────────────────────
function formatConfirmationMessage(data) {
  const items = data.items?.slice(0, 3).join(", ") || "N/A";
  return (
    `✅ *Gasto registrado exitosamente*\n\n` +
    `🏪 *Proveedor:* ${data.vendor || "N/A"}\n` +
    `📅 *Fecha:* ${data.date || "N/A"}\n` +
    `💰 *Total:* $${data.total?.toFixed(2)} ${data.currency}\n` +
    `🏷️ *Categoría:* ${data.category}\n` +
    `📁 *Proyecto:* ${data.project}\n` +
    `🛍️ *Concepto:* ${items}\n` +
    `🆔 *ID:* \`${data.id}\``
  );
}

// ─── Iniciar servidor ────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`🚀 Servidor corriendo en puerto ${PORT}`);
});
