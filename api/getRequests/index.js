/**
 * getRequests/index.js
 * GET /api/getRequests
 *
 * Görünürlük kuralları (sunucu tarafında zorunlu):
 *  - beklemede  → sadece talepçi + onaylayıcılar
 *  - onaylandi / reddedildi → talepçi kendi talebini, onaylayıcı hepsini görür
 *
 * Client'tan userId parametresi artık güvenlik için kullanılmıyor;
 * kimlik, JWT token'dan server-side çıkarılıyor.
 */

const { app } = require("@azure/functions");
const { getAllRequests } = require("../shared/excelService");
const { getAllApprovers } = require("../shared/excelService");
const { extractCaller, applyVisibilityRules } = require("../shared/authMiddleware");

app.http("getRequests", {
  methods: ["GET", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("getRequests çağrıldı");

    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
    };

    if (request.method === "OPTIONS") {
      return { status: 204, headers };
    }

    // ── 1. Kimliği token'dan çıkar ────────────────────────────────────────
    const caller = await extractCaller(request);
    if (!caller) {
      return {
        status: 401,
        headers,
        body: JSON.stringify({ error: "Kimlik doğrulama gerekli." }),
      };
    }

    try {
      // ── 2. Tüm talepleri ve onaylayıcı listesini paralel çek ──────────
      const [allRequests, approvers] = await Promise.all([
        getAllRequests(),
        getAllApprovers(),
      ]);

      const isApprover = approvers.some((a) => a.id === caller.userId);

      // ── 3. Görünürlük filtresi uygula ─────────────────────────────────
      const visible = applyVisibilityRules(allRequests, caller.userId, isApprover);

      // ── 4. Meta bilgi ekle (frontend'de UI ipuçları için) ─────────────
      const enriched = visible.map((req) => ({
        ...req,
        _visibility: getVisibilityMeta(req, caller.userId, isApprover),
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
      context.log.error("getRequests hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});

/**
 * Her talep için frontend'in göstereceği görünürlük metası
 */
function getVisibilityMeta(req, callerId, isApprover) {
  const isOwner = req.requesterId === callerId;

  if (req.status === "beklemede") {
    if (isOwner) {
      return {
        badge: "only_you_and_approvers",
        label: "Yalnızca siz ve onaylayıcılar görebilir",
        icon: "lock",
      };
    }
    if (isApprover) {
      return {
        badge: "approver_only",
        label: "Onaylayıcılara özel",
        icon: "lock",
      };
    }
  }

  return { badge: "normal", label: null, icon: null };
}
