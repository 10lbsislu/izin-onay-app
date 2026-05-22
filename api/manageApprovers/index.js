/**
 * manageApprovers/index.js
 * POST /api/manageApprovers
 * Body: { action: "add" | "remove", userId, displayName?, mail? }
 */

const { app } = require("@azure/functions");
const { addApprover, removeApprover } = require("../shared/excelService");

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
      const { action, userId, displayName, mail } = body;

      if (!action || !userId) {
        return {
          status: 400,
          headers,
          body: JSON.stringify({ error: "action ve userId zorunludur." }),
        };
      }

      if (action === "add") {
        const approver = {
          id: userId,
          displayName: displayName || "",
          mail: mail || "",
          addedAt: new Date().toISOString(),
        };
        await addApprover(approver);
        return {
          status: 201,
          headers,
          body: JSON.stringify({ approver }),
        };
      }

      if (action === "remove") {
        await removeApprover(userId);
        return {
          status: 200,
          headers,
          body: JSON.stringify({ removed: userId }),
        };
      }

      return {
        status: 400,
        headers,
        body: JSON.stringify({ error: "Bilinmeyen action." }),
      };
    } catch (err) {
      context.log.error("manageApprovers hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
