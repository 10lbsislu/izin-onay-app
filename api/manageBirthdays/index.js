/**
 * manageBirthdays/index.js
 * POST /api/manageBirthdays
 *
 * Doğum günü yönetimi (yalnızca tree admin):
 *   add    → { name, birthDate }  birthDate = "MM-DD"
 *   remove → { id }
 *
 * Listeleme getCalendar üzerinden (herkese açık) yapılır.
 */

const { app } = require("@azure/functions");
const { v4: uuidv4 } = require("uuid");
const { extractCaller } = require("../shared/authMiddleware");
const { addBirthday, removeBirthday, isTreeAdmin } = require("../shared/excelService");

// "MM-DD" doğrulama (01-12 / 01-31)
function isValidMMDD(s) {
  if (!/^\d{2}-\d{2}$/.test(s)) return false;
  const [mm, dd] = s.split("-").map(Number);
  return mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31;
}

app.http("manageBirthdays", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
    };

    if (request.method === "OPTIONS") return { status: 204, headers };

    const caller = await extractCaller(request);
    if (!caller) {
      return { status: 401, headers, body: JSON.stringify({ error: "Kimlik doğrulama gerekli." }) };
    }

    try {
      const callerIsAdmin = await isTreeAdmin(caller.userId);
      if (!callerIsAdmin) {
        return { status: 403, headers, body: JSON.stringify({ error: "Bu işlem için tree admin yetkisi gerekli." }) };
      }

      const body = await request.json();
      const { action } = body;

      if (action === "add") {
        const name = String(body.name || "").trim();
        const birthDate = String(body.birthDate || "").trim();
        if (!name) return { status: 400, headers, body: JSON.stringify({ error: "İsim zorunludur." }) };
        if (!isValidMMDD(birthDate)) {
          return { status: 400, headers, body: JSON.stringify({ error: "Geçersiz tarih (AA-GG formatı bekleniyor)." }) };
        }
        const birthday = { id: uuidv4(), name, birthDate };
        await addBirthday(birthday);
        return { status: 200, headers, body: JSON.stringify({ birthday }) };
      }

      if (action === "remove") {
        const { id } = body;
        if (!id) return { status: 400, headers, body: JSON.stringify({ error: "id zorunludur." }) };
        await removeBirthday(id);
        return { status: 200, headers, body: JSON.stringify({ removed: id }) };
      }

      return { status: 400, headers, body: JSON.stringify({ error: "Bilinmeyen action." }) };
    } catch (err) {
      context.error("manageBirthdays hata:", err);
      return { status: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  },
});
