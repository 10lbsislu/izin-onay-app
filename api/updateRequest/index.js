/**
 * updateRequest/index.js
 * PUT /api/updateRequest
 *
 * Güvenlik kuralları:
 *  - Sadece onaylayıcılar onayla/reddet yapabilir
 *  - Onaylayıcı kendi talebini onaylayamaz
 */

const { app } = require("@azure/functions");
const { updateRequest: updateInExcel, getAllApprovers } = require("../shared/excelService");
const { notifyUser } = require("../shared/notifyService");
const { extractCaller } = require("../shared/authMiddleware");

app.http("updateRequest", {
  methods: ["PUT", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("updateRequest çağrıldı");

    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "PUT, OPTIONS",
    };

    if (request.method === "OPTIONS") {
      return { status: 204, headers };
    }

    // ── Kimlik doğrulama ──────────────────────────────────────────────────
    const caller = await extractCaller(request);
    if (!caller) {
      return {
        status: 401,
        headers,
        body: JSON.stringify({ error: "Kimlik doğrulama gerekli." }),
      };
    }

    try {
      const body = await request.json();
      const { id, status, approverComment } = body;

      if (!id || !status) {
        return {
          status: 400,
          headers,
          body: JSON.stringify({ error: "id ve status zorunludur." }),
        };
      }

      if (!["onaylandi", "reddedildi"].includes(status)) {
        return {
          status: 400,
          headers,
          body: JSON.stringify({ error: "Geçersiz durum değeri." }),
        };
      }

      // ── Onaylayıcı yetkisi kontrolü ───────────────────────────────────
      const approvers = await getAllApprovers();
      const isApprover = approvers.some((a) => a.id === caller.userId);

      if (!isApprover) {
        return {
          status: 403,
          headers,
          body: JSON.stringify({
            error: "Bu işlem için onaylayıcı yetkisi gerekiyor.",
          }),
        };
      }

      // ── Kendi talebini onaylayamaz ─────────────────────────────────────
      // (requesterId talebin sahibi — excel'den kontrol edelim)
      const { getAllRequests } = require("../shared/excelService");
      const allRequests = await getAllRequests();
      const targetRequest = allRequests.find((r) => r.id === id);

      if (!targetRequest) {
        return {
          status: 404,
          headers,
          body: JSON.stringify({ error: "Talep bulunamadı." }),
        };
      }

      if (targetRequest.requesterId === caller.userId) {
        return {
          status: 403,
          headers,
          body: JSON.stringify({
            error: "Kendi izin talebinizi onaylayamazsınız.",
          }),
        };
      }

      if (targetRequest.status !== "beklemede") {
        return {
          status: 409,
          headers,
          body: JSON.stringify({
            error: "Bu talep zaten işlem görmüş: " + targetRequest.status,
          }),
        };
      }

      // ── Excel güncelle ────────────────────────────────────────────────
      const updatedRequest = await updateInExcel(id, {
        status,
        approverComment: approverComment || "",
        approverId:      caller.userId,
        approverName:    caller.name || "",
        updatedAt:       new Date().toISOString(),
      });

      // ── Çalışana kişisel bildirim ─────────────────────────────────────
      notifyUser(updatedRequest.requesterId, updatedRequest).catch((e) =>
        context.log.warn("Kullanıcı bildirimi başarısız:", e.message)
      );

      return {
        status: 200,
        headers,
        body: JSON.stringify({ request: updatedRequest }),
      };
    } catch (err) {
      context.log.error("updateRequest hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
