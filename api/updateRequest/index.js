/**
 * updateRequest/index.js
 * PUT /api/updateRequest
 *
 * Güvenlik kuralları:
 *  - Sadece onaylayıcılar onayla/reddet yapabilir
 *  - Onaylayıcı kendi talebini onaylayamaz
 */

const { app } = require("@azure/functions");
const { updateRequest: updateInExcel, getAllRequests, isTreeAdmin, getApprovalViewers } = require("../shared/excelService");
const { notifyUser, notifyApprovalViewers } = require("../shared/notifyService");
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

      // ── Talebi bul ────────────────────────────────────────────────────
      const allRequests = await getAllRequests();
      const targetRequest = allRequests.find((r) => r.id === id);

      if (!targetRequest) {
        return {
          status: 404,
          headers,
          body: JSON.stringify({ error: "Talep bulunamadı." }),
        };
      }

      // ── Hiyerarşik yetki: atanmış amir veya (yetim talepler için) tree admin ──
      const isAssignedApprover = targetRequest.approverId === caller.userId;
      const isOrphan = !targetRequest.approverId;
      let isFallbackAdmin = false;
      if (!isAssignedApprover && isOrphan) {
        // Eski/yetim talep: tree admin müdahale edebilir
        isFallbackAdmin = await isTreeAdmin(caller.userId);
      }

      if (!isAssignedApprover && !isFallbackAdmin) {
        const detail = isOrphan
          ? "Bu talebin atanmış amiri yok (büyük olasılıkla eski bir kayıt). Yalnızca tree admin temizleyebilir."
          : `Bu talebi yalnızca atanmış amir onaylayabilir${targetRequest.approverName ? ` (${targetRequest.approverName})` : ""}.`;
        return {
          status: 403,
          headers,
          body: JSON.stringify({ error: detail }),
        };
      }

      // ── Defansif: kendi talebini onaylayamaz ──────────────────────────
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

      // ── Excel güncelle (approverId/Name createRequest'te yazıldı, değişmez) ──
      // Tree admin yetim talebi onaylıyorsa approverId'yi kendisine yaz (audit için)
      const adminFallbackPatch = isFallbackAdmin ? {
        approverId: caller.userId,
        approverName: caller.name || "",
      } : {};

      const updatedRequest = await updateInExcel(id, {
        ...adminFallbackPatch,
        status,
        approverComment: approverComment || "",
        updatedAt:       new Date().toISOString(),
      });

      // ── Çalışana kişisel bildirim ─────────────────────────────────────
      notifyUser(updatedRequest.requesterId, updatedRequest).catch((e) =>
        context.log.warn("Kullanıcı bildirimi başarısız:", e.message)
      );

      // ── Onaylandıysa: "onaylanan izin görüntüleyici"lere bildirim ─────
      if (status === "onaylandi") {
        getApprovalViewers()
          .then((viewers) =>
            notifyApprovalViewers(updatedRequest, viewers, caller.email)
          )
          .catch((e) =>
            context.log.warn("Görüntüleyici bildirimi başarısız:", e.message)
          );
      }

      return {
        status: 200,
        headers,
        body: JSON.stringify({ request: updatedRequest }),
      };
    } catch (err) {
      context.error("updateRequest hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
