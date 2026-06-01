/**
 * createRequest/index.js
 * POST /api/createRequest
 *
 * Talep oluşturma — kimlik token'dan alınır, body'den gelen requesterId
 * token ile çapraz kontrol edilir (spoofing önleme).
 */

const { app } = require("@azure/functions");
const { v4: uuidv4 } = require("uuid");
const { addRequest, getHierarchyNode } = require("../shared/excelService");
const { notifyChannel, notifyApproverByEmail, notifyApproverInTeams } = require("../shared/notifyService");
const { extractCaller } = require("../shared/authMiddleware");

app.http("createRequest", {
  methods: ["POST", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("createRequest çağrıldı");

    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
    };

    if (request.method === "OPTIONS") {
      return { status: 204, headers };
    }

    // Kimlik doğrulama
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

      // Güvenlik: requesterId token'daki kimlikle eşleşmeli
      if (body.requesterId && body.requesterId !== caller.userId) {
        return {
          status: 403,
          headers,
          body: JSON.stringify({
            error: "Başkası adına talep oluşturamazsınız.",
          }),
        };
      }

      // Açıklama zorunlu
      if (!body.description || !String(body.description).trim()) {
        return {
          status: 400,
          headers,
          body: JSON.stringify({ error: "Açıklama zorunludur." }),
        };
      }

      // Hiyerarşi: kullanıcının amirini bul, talebi ona ata
      const callerNode = await getHierarchyNode(caller.userId);
      const callerIsTreeAdmin = callerNode &&
        String(callerNode.isTreeAdmin).toLowerCase() === "true";

      let approverId = "";
      let approverName = "";
      let approverEmail = "";

      if (callerNode && callerNode.managerId) {
        // Normal kullanıcı: amir otomatik
        approverId = callerNode.managerId;
        const managerNode = await getHierarchyNode(approverId);
        approverName = managerNode ? managerNode.displayName : "";
        approverEmail = managerNode ? managerNode.mail : "";
      } else if (callerIsTreeAdmin) {
        // Root/tree admin: payload'dan approverId al
        if (!body.approverId) {
          return {
            status: 400,
            headers,
            body: JSON.stringify({
              error: "Onaylayıcı seçmelisiniz (kendiniz hiyerarşinin tepesindesiniz).",
            }),
          };
        }
        const targetNode = await getHierarchyNode(body.approverId);
        if (!targetNode) {
          return {
            status: 400,
            headers,
            body: JSON.stringify({ error: "Geçersiz onaylayıcı seçimi." }),
          };
        }
        if (targetNode.id === caller.userId) {
          return {
            status: 400,
            headers,
            body: JSON.stringify({ error: "Kendinizi onaylayıcı seçemezsiniz." }),
          };
        }
        approverId = targetNode.id;
        approverName = targetNode.displayName;
        approverEmail = targetNode.mail;
      } else {
        return {
          status: 400,
          headers,
          body: JSON.stringify({
            error: "Hiyerarşide tanımlı değilsiniz. Yöneticinizle iletişime geçin.",
          }),
        };
      }

      const newRequest = {
        id: uuidv4(),
        requesterId:    caller.userId,
        requesterName:  caller.name  || body.requesterName  || "",
        requesterEmail: caller.email || body.requesterEmail || "",
        leaveType:      body.leaveType,
        startDate:      body.startDate,
        endDate:        body.endDate,
        startTime:      body.startTime || "",
        endTime:        body.endTime || "",
        totalDays:      Number(body.totalDays),
        description:    body.description || "",
        status:         "beklemede",
        approverId,
        approverName,
        approverComment: "",
        createdAt:      new Date().toISOString(),
        updatedAt:      "",
      };

      await addRequest(newRequest);

      // Kanala Adaptive Card bildirimi (hata ana akışı durdurmaz)
      notifyChannel(newRequest).catch((e) =>
        context.log.warn("Kanal bildirimi başarısız:", e.message)
      );

      // Onaylayıcıya e-posta bildirimi
      notifyApproverByEmail(newRequest, approverEmail).catch((e) =>
        context.log.warn("Mail bildirimi başarısız:", e.message)
      );

      // Onaylayıcıya Teams chat bildirimi
      notifyApproverInTeams(newRequest, approverId).catch((e) =>
        context.log.warn("Teams bildirimi başarısız:", e.message)
      );

      return {
        status: 201,
        headers,
        body: JSON.stringify({ request: newRequest }),
      };
    } catch (err) {
      context.error("createRequest hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
