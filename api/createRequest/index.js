/**
 * createRequest/index.js
 * POST /api/createRequest
 *
 * Talep oluşturma — kimlik token'dan alınır, body'den gelen requesterId
 * token ile çapraz kontrol edilir (spoofing önleme).
 */

const { app } = require("@azure/functions");
const { v4: uuidv4 } = require("uuid");
const { addRequest } = require("../shared/excelService");
const { notifyChannel } = require("../shared/notifyService");
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

      const newRequest = {
        id: uuidv4(),
        requesterId:    caller.userId,
        requesterName:  caller.name  || body.requesterName  || "",
        requesterEmail: caller.email || body.requesterEmail || "",
        leaveType:      body.leaveType,
        startDate:      body.startDate,
        endDate:        body.endDate,
        totalDays:      Number(body.totalDays),
        description:    body.description || "",
        status:         "beklemede",
        approverId:     "",
        approverName:   "",
        approverComment: "",
        createdAt:      new Date().toISOString(),
        updatedAt:      "",
      };

      await addRequest(newRequest);

      // Kanala Adaptive Card bildirimi (hata ana akışı durdurmaz)
      notifyChannel(newRequest).catch((e) =>
        context.log.warn("Kanal bildirimi başarısız:", e.message)
      );

      return {
        status: 201,
        headers,
        body: JSON.stringify({ request: newRequest }),
      };
    } catch (err) {
      context.log.error("createRequest hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
