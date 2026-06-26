/**
 * getCalendar/index.js
 * GET /api/getCalendar
 *
 * Takvim için HERKESE AÇIK (kimliği doğrulanmış her kullanıcı) veri döndürür:
 *  - leaves:    tüm ONAYLANAN izinler (sadeleştirilmiş alanlar — gizlilik için)
 *  - birthdays: tüm doğum günleri
 */

const { app } = require("@azure/functions");
const {
  getAllRequests, getAllBirthdays, getDirectReports, isTreeAdmin, isApprovalViewer,
} = require("../shared/excelService");
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
      const [allRequests, birthdays, directReports, treeAdmin, viewer] = await Promise.all([
        getAllRequests(),
        getAllBirthdays(),
        getDirectReports(caller.userId),
        isTreeAdmin(caller.userId),
        isApprovalViewer(caller.userId),
      ]);

      // Yönetici: astı olan VEYA tree admin VEYA onaylanan-izin görüntüleyici
      const isManager = directReports.length > 0 || treeAdmin || viewer;

      // Görünürlük kuralı:
      //   - yıllık izinler → herkese açık
      //   - saatlik/ücretsiz/diğer → yalnızca yöneticiler
      const leaves = allRequests
        .filter((r) => r.status === "onaylandi")
        .filter((r) => isManager || r.leaveType === "yillik")
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
        body: JSON.stringify({ leaves, birthdays, meta: { isManager } }),
      };
    } catch (err) {
      context.error("getCalendar hata:", err);
      return { status: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  },
});
