/**
 * searchUsers/index.js
 * GET /api/searchUsers?q=arama
 * → Graph API /users?$search="displayName:{q}" ile kuruluş kullanıcıları arar
 * Bu, AdminPanel'deki UserPicker bileşenini besler.
 */

const { app } = require("@azure/functions");
const { getAppGraphClient } = require("../shared/graphClient");

app.http("searchUsers", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    context.log("searchUsers çağrıldı");

    const headers = {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
    };

    try {
      const q = request.query.get("q");
      if (!q || q.trim().length < 2) {
        return {
          status: 400,
          headers,
          body: JSON.stringify({ error: "En az 2 karakter girin." }),
        };
      }

      const client = getAppGraphClient();

      // Graph kullanıcı arama — ConsistencyLevel: eventual gerektirir
      const result = await client
        .api("/users")
        .header("ConsistencyLevel", "eventual")
        .search(`"displayName:${q}"`)
        .select("id,displayName,mail,jobTitle,department,userPrincipalName")
        .top(10)
        .orderby("displayName")
        .filter("accountEnabled eq true")
        .get();

      const users = (result.value || []).map((u) => ({
        id: u.id,
        displayName: u.displayName,
        mail: u.mail || u.userPrincipalName,
        jobTitle: u.jobTitle,
        department: u.department,
        userPrincipalName: u.userPrincipalName,
      }));

      return {
        status: 200,
        headers,
        body: JSON.stringify({ users }),
      };
    } catch (err) {
      context.log.error("searchUsers hata:", err);
      return {
        status: 500,
        headers,
        body: JSON.stringify({ error: err.message }),
      };
    }
  },
});
