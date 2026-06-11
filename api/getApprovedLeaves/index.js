/**
 * getApprovedLeaves/index.js
 * GET /api/getApprovedLeaves
 *
 * Yalnızca "onaylanan izin görüntüleyici" (isApprovalViewer) veya tree admin
 * yetkisine sahip kullanıcılar ÇAĞIRABİLİR. Kuruluştaki TÜM onaylanan
 * izinleri döndürür (hiyerarşik görünürlük kuralı uygulanmaz).
 */

const { app } = require("@azure/functions");
const { getAllRequests, isApprovalViewer, isTreeAdmin } = require("../shared/excelService");
const { extractCaller } = require("../shared/authMiddleware");

app.http("getApprovedLeaves", {
  methods: ["GET", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
    };

    if (request.method === "OPTIONS") {
      return { status: 204, headers };
    }

    const caller = await extractCaller(request);
    if (!caller) {
      return {
        status: 401,
        headers,
        body: JSON.stringify({ error: "Kimlik doğrulama gerekli." }),
      };
    }

    try {
      // Yetki: onaylanan izin görüntüleyici VEYA tree admin
      const [viewer, admin] = await Promise.all([
        isApprovalViewer(caller.userId),
        isTreeAdmin(caller.userId),
      ]);

      if (!viewer && !admin) {
        return {
          status: 403,
          headers,
          body: JSON.stringify({
            error: "Bu sekmeyi görüntüleme yetkiniz yok.",
          }),
        };
      }

      const all = await getAllRequests();
      const approved = all
        .filter((r) => r.status === "onaylandi")
        .sort((a, b) => String(b.updatedAt || b.createdAt).localeCompare(String(a.updatedAt || a.createdAt)));

      return {
        status: 200,
        headers,
        body: JSON.stringify({
          requests: approved,
          meta: {
            callerId: caller.userId,
            total: approved.length,
          },
        }),
      };
    } catch (err) {
      context.error("getApprovedLeaves hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
