/**
 * getApprovers/index.js
 * GET /api/getApprovers → Excel Onaylayıcılar tablosunu döner
 */

const { app } = require("@azure/functions");
const { getAllApprovers } = require("../shared/excelService");

app.http("getApprovers", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
    };

    try {
      const approvers = await getAllApprovers();
      return {
        status: 200,
        headers,
        body: JSON.stringify({ approvers }),
      };
    } catch (err) {
      context.log.error("getApprovers hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
