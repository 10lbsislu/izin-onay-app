/**
 * getRequests/index.js
 * GET /api/getRequests
 *
 * Hiyerarşik görünürlük:
 *  - Kullanıcı kendi taleplerini her zaman görür
 *  - Bir kullanıcı YALNIZCA doğrudan astlarının taleplerini görür
 *  - Zincirde yukarı yayılma YOK
 */

const { app } = require("@azure/functions");
const { getAllRequests, getDirectReports } = require("../shared/excelService");
const { extractCaller, applyVisibilityRules } = require("../shared/authMiddleware");

app.http("getRequests", {
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
      const [allRequests, directReports] = await Promise.all([
        getAllRequests(),
        getDirectReports(caller.userId),
      ]);

      const directReportIds = new Set(directReports.map((n) => n.id));
      const isApprover = directReportIds.size > 0;

      const visible = applyVisibilityRules(allRequests, caller.userId, directReportIds);

      const enriched = visible.map((req) => ({
        ...req,
        _visibility: getVisibilityMeta(req, caller.userId),
      }));

      return {
        status: 200,
        headers,
        body: JSON.stringify({
          requests: enriched,
          meta: {
            callerId: caller.userId,
            isApprover,
            total: enriched.length,
          },
        }),
      };
    } catch (err) {
      context.error("getRequests hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});

function getVisibilityMeta(req, callerId) {
  const isOwner = req.requesterId === callerId;

  if (req.status === "beklemede") {
    if (isOwner) {
      return {
        badge: "only_you_and_approvers",
        label: "Yalnızca siz ve amiriniz görebilir",
        icon: "lock",
      };
    }
    // Bu noktada caller bu talebi görüyorsa caller, talepçinin amiridir
    return {
      badge: "approver_only",
      label: "Sadece amir olarak siz görüyorsunuz",
      icon: "lock",
    };
  }

  return { badge: "normal", label: null, icon: null };
}
