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
          {
            title: "Tarihler",
            value: `${formatDate(request.startDate)} — ${formatDate(request.endDate)}`,
          },
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
          {
            title: "Tarihler",
            value: `${formatDate(request.startDate)} — ${formatDate(request.endDate)}`,
          },
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
    hastalik: "Hastalık İzni",
    mazeret: "Mazeret İzni",
    ucretsiz: "Ücretsiz İzin",
    diger: "Diğer",
  };
  return labels[type] || type;
}

function formatDate(dateStr) {
  if (!dateStr) return "—";
  return new Date(dateStr).toLocaleDateString("tr-TR");
}

module.exports = { notifyChannel, notifyUser };
