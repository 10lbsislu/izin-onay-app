/**
 * manageApprovers/index.js
 * POST /api/manageApprovers
 *
 * Hiyerarşi yönetimi action'ları:
 *   get-tree      → tüm hiyerarşi node'larını döndür (herkes)
 *   set-manager   → kullanıcının amirini ata/değiştir (tree admin)
 *   remove        → hiyerarşi node'unu sil (tree admin, çocuk yoksa)
 *   grant-admin   → tree admin yetkisi ver (tree admin)
 *   revoke-admin  → tree admin yetkisini kaldır (tree admin, son admin değilse)
 */

const { app } = require("@azure/functions");
const { extractCaller } = require("../shared/authMiddleware");
const {
  getAllApprovers,
  getHierarchyNode,
  setUserManager,
  setTreeAdmin,
  removeHierarchyNode,
  isTreeAdmin,
  countTreeAdmins,
  setApprovalViewer,
} = require("../shared/excelService");

app.http("manageApprovers", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
    };

    try {
      const body = await request.json();
      const { action } = body;

      if (!action) {
        return { status: 400, headers, body: JSON.stringify({ error: "action zorunludur." }) };
      }

      // get-tree herkese açık
      if (action === "get-tree") {
        const nodes = await getAllApprovers();
        const normalized = nodes.map((n) => ({
          ...n,
          isTreeAdmin: String(n.isTreeAdmin).toLowerCase() === "true",
          isApprovalViewer: String(n.isApprovalViewer).toLowerCase() === "true",
          managerId: n.managerId || "",
        }));
        return { status: 200, headers, body: JSON.stringify({ nodes: normalized }) };
      }

      // Yazma action'ları için yetki kontrolü
      const caller = await extractCaller(request);
      if (!caller) {
        return { status: 401, headers, body: JSON.stringify({ error: "Yetkilendirme başarısız." }) };
      }

      const callerIsAdmin = await isTreeAdmin(caller.userId);
      if (!callerIsAdmin) {
        return { status: 403, headers, body: JSON.stringify({ error: "Bu işlem için tree admin yetkisi gerekli." }) };
      }

      if (action === "set-manager") {
        const { userId, managerId, displayName, mail } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        const node = await setUserManager(userId, managerId || "", displayName, mail);
        return { status: 200, headers, body: JSON.stringify({ node }) };
      }

      if (action === "remove") {
        const { userId } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        await removeHierarchyNode(userId);
        return { status: 200, headers, body: JSON.stringify({ removed: userId }) };
      }

      if (action === "grant-admin") {
        const { userId } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        const target = await getHierarchyNode(userId);
        if (!target) {
          return { status: 404, headers, body: JSON.stringify({ error: "Kullanıcı hiyerarşide yok." }) };
        }
        const node = await setTreeAdmin(userId, true);
        return { status: 200, headers, body: JSON.stringify({ node }) };
      }

      if (action === "revoke-admin") {
        const { userId } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        const adminCount = await countTreeAdmins();
        if (adminCount <= 1) {
          return { status: 400, headers, body: JSON.stringify({ error: "Son tree admin'i kaldıramazsınız." }) };
        }
        const node = await setTreeAdmin(userId, false);
        return { status: 200, headers, body: JSON.stringify({ node }) };
      }

      if (action === "grant-viewer") {
        const { userId } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        const target = await getHierarchyNode(userId);
        if (!target) {
          return { status: 404, headers, body: JSON.stringify({ error: "Kullanıcı hiyerarşide yok." }) };
        }
        const node = await setApprovalViewer(userId, true);
        return { status: 200, headers, body: JSON.stringify({ node }) };
      }

      if (action === "revoke-viewer") {
        const { userId } = body;
        if (!userId) {
          return { status: 400, headers, body: JSON.stringify({ error: "userId zorunludur." }) };
        }
        const node = await setApprovalViewer(userId, false);
        return { status: 200, headers, body: JSON.stringify({ node }) };
      }

      return { status: 400, headers, body: JSON.stringify({ error: "Bilinmeyen action." }) };
    } catch (err) {
      context.error("manageApprovers hata:", err);
      return { status: 500, headers, body: JSON.stringify({ error: err.message }) };
    }
  },
});
