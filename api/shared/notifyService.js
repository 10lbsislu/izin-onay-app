/**
 * shared/notifyService.js
 * Teams bildirimleri:
 *  1. Kanala Adaptive Card (yeni talep açıldığında onaylayıcılara)
 *  2. Kişiye direkt mesaj (talep sonucu çalışana)
 */

const { getAppGraphClient } = require("./graphClient");

const TEAM_ID    = process.env.TEAMS_TEAM_ID;
const CHANNEL_ID = process.env.TEAMS_CHANNEL_ID;

// ─── Adaptive Card şablonları ─────────────────────────────────────────────────

function newRequestCard(request, appUrl) {
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "📋 Yeni İzin Talebi",
        weight: "Bolder",
        size: "Large",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Çalışan", value: request.requesterName },
          { title: "İzin Türü", value: izinTuruLabel(request.leaveType) },
          { title: "Tarihler", value: formatRequestDateRange(request) },
          { title: "Süre", value: `${request.totalDays} iş günü` },
          ...(request.description
            ? [{ title: "Açıklama", value: request.description }]
            : []),
        ],
      },
      {
        type: "TextBlock",
        text: `Talep Zamanı: ${new Date(request.createdAt).toLocaleString("tr-TR")}`,
        size: "Small",
        color: "Accent",
        isSubtle: true,
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "🔍 Talebi İncele ve Onayla",
        url: appUrl,
        style: "positive",
      },
    ],
  };
}

function requestResultCard(request) {
  const isApproved = request.status === "onaylandi";
  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: isApproved ? "✅ İzin Talebiniz Onaylandı" : "❌ İzin Talebiniz Reddedildi",
        weight: "Bolder",
        size: "Large",
        color: isApproved ? "Good" : "Attention",
      },
      {
        type: "FactSet",
        facts: [
          { title: "İzin Türü", value: izinTuruLabel(request.leaveType) },
          { title: "Tarihler", value: formatRequestDateRange(request) },
          { title: "Süre", value: `${request.totalDays} iş günü` },
          { title: "İşlem Yapan", value: request.approverName || "—" },
          ...(request.approverComment
            ? [{ title: "Yorum", value: request.approverComment }]
            : []),
        ],
      },
    ],
  };
}

// ─── Kanal mesajı ─────────────────────────────────────────────────────────────

/**
 * Belirlenen Teams kanalına yeni talep Adaptive Card'ı gönder.
 */
async function notifyChannel(request) {
  const client = getAppGraphClient();
  const appUrl = process.env.FRONTEND_URL || "https://teams.microsoft.com";

  try {
    await client
      .api(`/teams/${TEAM_ID}/channels/${CHANNEL_ID}/messages`)
      .post({
        subject: `İzin Talebi: ${request.requesterName}`,
        body: {
          contentType: "html",
          content: "<attachment id=\"card1\"></attachment>",
        },
        attachments: [
          {
            id: "card1",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: JSON.stringify(newRequestCard(request, appUrl)),
          },
        ],
      });
    console.log("Kanal bildirimi gönderildi.");
  } catch (err) {
    console.error("Kanal bildirimi hatası:", err.message);
    // Bildirim hataları ana akışı durdurmasın
  }
}

/**
 * Çalışana kişisel chat mesajı gönder (talep sonucu).
 */
async function notifyUser(userId, request) {
  const client = getAppGraphClient();

  try {
    // 1. Bot ile çalışan arasında chat bul veya oluştur
    const chat = await client
      .api("/chats")
      .post({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${process.env.BOT_USER_ID || userId}`,
          },
        ],
      })
      .catch(async () => {
        // Chat zaten varsa listeden bul
        const chats = await client.api(`/users/${userId}/chats`).get();
        return chats.value?.[0];
      });

    if (!chat?.id) throw new Error("Chat ID alınamadı.");

    // 2. Adaptive Card mesajı gönder
    await client
      .api(`/chats/${chat.id}/messages`)
      .post({
        body: {
          contentType: "html",
          content: "<attachment id=\"result1\"></attachment>",
        },
        attachments: [
          {
            id: "result1",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: JSON.stringify(requestResultCard(request)),
          },
        ],
      });

    console.log(`Kullanıcı bildirimi gönderildi: ${userId}`);
  } catch (err) {
    console.error("Kullanıcı bildirimi hatası:", err.message);
  }
}

// ─── Yardımcılar ─────────────────────────────────────────────────────────────

function izinTuruLabel(type) {
  const labels = {
    yillik: "Yıllık İzin",
    saatlik: "Saatlik İzin",
    ucretsiz: "Ücretsiz İzin",
    diger: "Diğer",
  };
  return labels[type] || type;
}

function formatDate(dateStr) {
  if (!dateStr) return "—";
  return new Date(dateStr).toLocaleDateString("tr-TR");
}

function formatRequestDateRange(request) {
  if (request.leaveType === "saatlik" && request.startTime && request.endTime) {
    return `${formatDate(request.startDate)} ${request.startTime}–${request.endTime}`;
  }
  if (request.startDate === request.endDate) {
    return formatDate(request.startDate);
  }
  return `${formatDate(request.startDate)} — ${formatDate(request.endDate)}`;
}

// ─── Onaylayıcıya mail bildirimi ─────────────────────────────────────────────

/**
 * Teams deep link oluşturur. TEAMS_APP_ID env değişkeni Teams Admin Center'a
 * yüklenen app'in catalog ID'sidir (manifest.json'daki id ile aynı).
 * Env yoksa web URL'ye düşer.
 */
function buildTeamsDeepLink() {
  const appId = process.env.TEAMS_APP_ID;
  const entityId = process.env.TEAMS_ENTITY_ID || "izin-onay-tab";
  if (!appId) return null;
  return `https://teams.microsoft.com/l/entity/${appId}/${entityId}`;
}

async function notifyApproverByEmail(request, approverEmail) {
  if (!approverEmail) {
    console.warn("notifyApproverByEmail: onaylayıcı maili yok, atlandı.");
    return;
  }
  const client = getAppGraphClient();
  const webUrl = process.env.FRONTEND_URL || "https://teams.microsoft.com";
  const teamsUrl = buildTeamsDeepLink();
  const subject = `İzin Talebi — ${request.requesterName}`;
  const dateLine = formatRequestDateRange(request);
  const sureLine = request.leaveType === "saatlik"
    ? `${request.totalDays} iş günü karşılığı (saatlik)`
    : `${request.totalDays} iş günü`;

  const primaryUrl = teamsUrl || webUrl;
  const primaryLabel = teamsUrl ? "🔍 Teams'te Aç ve Onayla" : "🔍 Talebi İncele ve Onayla";

  const buttonsHtml = teamsUrl
    ? `
      <p style="margin-top: 16px;">
        <a href="${teamsUrl}" style="background:#5b5fc7;color:white;padding:10px 20px;border-radius:4px;text-decoration:none;display:inline-block;margin-right:8px;">
          🔍 Teams'te Aç ve Onayla
        </a>
        <a href="${webUrl}" style="background:#0078d4;color:white;padding:10px 20px;border-radius:4px;text-decoration:none;display:inline-block;">
          🌐 Web'de Aç
        </a>
      </p>
      <p style="color:#666; font-size:11px; margin-top:8px;">
        Teams butonu Teams uygulamasında bir kez daha giriş yapmadan açar.
      </p>`
    : `
      <p style="margin-top: 16px;">
        <a href="${primaryUrl}" style="background:#0078d4;color:white;padding:10px 20px;border-radius:4px;text-decoration:none;display:inline-block;">
          ${primaryLabel}
        </a>
      </p>`;

  const html = `
    <div style="font-family: Segoe UI, Arial, sans-serif; max-width: 560px;">
      <h2 style="color:#0078d4; margin-bottom: 8px;">📋 Yeni İzin Talebi</h2>
      <p><strong>${request.requesterName}</strong> adlı çalışan size bir izin talebi gönderdi.</p>
      <table style="border-collapse: collapse; margin: 12px 0;">
        <tr><td style="padding:4px 12px 4px 0;"><b>İzin Türü:</b></td><td>${izinTuruLabel(request.leaveType)}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Tarih:</b></td><td>${dateLine}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Süre:</b></td><td>${sureLine}</td></tr>
        ${request.description ? `<tr><td style="padding:4px 12px 4px 0;vertical-align:top;"><b>Açıklama:</b></td><td>${request.description}</td></tr>` : ""}
        <tr><td style="padding:4px 12px 4px 0;"><b>E-posta:</b></td><td>${request.requesterEmail || "—"}</td></tr>
      </table>
      ${buttonsHtml}
      <p style="color:#888; font-size:12px; margin-top:24px;">
        Bu e-posta İzin Onay Sistemi tarafından otomatik gönderildi.
      </p>
    </div>
  `;

  // Mail'i talepçinin kutusundan gönder — alıcı kim olduğunu kolayca görsün
  const senderUserId = request.requesterEmail || request.requesterId;

  console.log(`[MAIL] gönderiliyor: from=${senderUserId} to=${approverEmail}`);

  try {
    await client
      .api(`/users/${senderUserId}/sendMail`)
      .post({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: [{ emailAddress: { address: approverEmail } }],
        },
        saveToSentItems: true,
      });
    console.log(`[MAIL] BAŞARILI: ${approverEmail}`);
  } catch (err) {
    console.error(`[MAIL] BAŞARISIZ from=${senderUserId} to=${approverEmail}:`, {
      message: err.message,
      statusCode: err.statusCode,
      code: err.code,
      body: typeof err.body === "string" ? err.body : JSON.stringify(err.body),
    });
  }
}

// ─── Onaylayıcıya Teams chat bildirimi ───────────────────────────────────────

/**
 * Talepçi ile onaylayıcı arasında 1-1 Teams chat'i bul/oluştur,
 * Adaptive Card mesajı gönder. Hata olursa ana akışı durdurmaz.
 */
async function notifyApproverInTeams(request, approverId) {
  if (!approverId) return;
  const client = getAppGraphClient();
  const appUrl = process.env.FRONTEND_URL || "https://teams.microsoft.com";
  const teamsUrl = buildTeamsDeepLink() || appUrl;

  const card = {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "📋 Yeni İzin Talebi",
        weight: "Bolder",
        size: "Large",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: `${request.requesterName} sizden onay bekliyor.`,
        wrap: true,
      },
      {
        type: "FactSet",
        facts: [
          { title: "İzin Türü", value: izinTuruLabel(request.leaveType) },
          { title: "Tarih", value: formatRequestDateRange(request) },
          { title: "Süre", value: request.leaveType === "saatlik"
              ? `${request.totalDays} iş günü karşılığı (saatlik)`
              : `${request.totalDays} iş günü` },
          ...(request.description ? [{ title: "Açıklama", value: request.description }] : []),
        ],
      },
    ],
    actions: [
      {
        type: "Action.OpenUrl",
        title: "🔍 Talebi İncele ve Onayla",
        url: teamsUrl,
        style: "positive",
      },
    ],
  };

  console.log(`[TEAMS] chat oluşturuluyor: requester=${request.requesterId} approver=${approverId}`);

  let chat;
  try {
    // 1) Talepçi ↔ Onaylayıcı arasında 1-1 chat oluştur (varsa Graph aynısını döner)
    chat = await client
      .api("/chats")
      .post({
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${request.requesterId}`,
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${approverId}`,
          },
        ],
      });

    if (!chat?.id) throw new Error("Chat ID alınamadı.");
    console.log(`[TEAMS] chat OK: ${chat.id}`);
  } catch (err) {
    console.error(`[TEAMS] chat oluşturma BAŞARISIZ:`, {
      message: err.message,
      statusCode: err.statusCode,
      code: err.code,
      body: typeof err.body === "string" ? err.body : JSON.stringify(err.body),
    });
    return;
  }

  try {
    // 2) Adaptive Card mesajı gönder
    await client
      .api(`/chats/${chat.id}/messages`)
      .post({
        body: {
          contentType: "html",
          content: '<attachment id="leaveCard"></attachment>',
        },
        attachments: [
          {
            id: "leaveCard",
            contentType: "application/vnd.microsoft.card.adaptive",
            content: JSON.stringify(card),
          },
        ],
      });

    console.log(`[TEAMS] mesaj BAŞARILI: ${approverId}`);
  } catch (err) {
    console.error(`[TEAMS] mesaj gönderme BAŞARISIZ:`, {
      message: err.message,
      statusCode: err.statusCode,
      code: err.code,
      body: typeof err.body === "string" ? err.body : JSON.stringify(err.body),
    });
  }
}

module.exports = { notifyChannel, notifyUser, notifyApproverByEmail, notifyApproverInTeams };
