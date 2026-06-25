/**
 * getCalendar/index.js
 * GET /api/getCalendar
 *
 * Takvim için HERKESE AÇIK (kimliği doğrulanmış her kullanıcı) veri döndürür:
 *  - leaves:    tüm ONAYLANAN izinler (sadeleştirilmiş alanlar — gizlilik için)
 *  - birthdays: tüm doğum günleri
 */

const { app } = require("@azure/functions");
const { getAllRequests, getAllBirthdays } = require("../shared/excelService");
const { extractCaller } = require("../shared/authMiddleware");

app.http("getCalendar", {
  methods: ["GET", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
    };

    if (request.method === "OPTIONS") return { status: 204, headers };

    const caller = await extractCaller(request);
    if (!caller) {
      return { status: 401, headers, body: JSON.stringify({ error: "Kimlik doğrulama gerekli." }) };
    }

    try {
      const [allRequests, birthdays] = await Promise.all([
        getAllRequests(),
        getAllBirthdays(),
      ]);

      // Sadece onaylanan izinler; gizlilik için sınırlı alanlar (açıklama yok)
      const leaves = allRequests
        .filter((r) => r.status === "onaylandi")
        .map((r) => ({
          id: r.id,
          requesterName: r.requesterName,
          leaveType: r.leaveType,
          startDate: r.startDate,
          endDate: r.endDate,
          startTime: r.startTime || "",
          endTime: r.endTime || "",
        }));

      return {
        status: 200,
        headers,
        body: JSON.stringify({ leaves, birthdays }),
      };
    } catch (err) {
      context.error("getCalendar hata:", err);
      return { status: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  },
});
