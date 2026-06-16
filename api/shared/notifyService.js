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

async function notifyApproverByEmail(request, approverEmail, opts = {}) {
  if (!approverEmail) {
    console.warn("notifyApproverByEmail: onaylayıcı maili yok, atlandı.");
    return;
  }
  const reminder = !!opts.reminder;
  const client = getAppGraphClient();
  const webUrl = process.env.FRONTEND_URL || "https://teams.microsoft.com";
  const teamsUrl = buildTeamsDeepLink();
  const subject = reminder
    ? `⏰ Hatırlatma: Bekleyen İzin Talebi — ${request.requesterName}`
    : `İzin Talebi — ${request.requesterName}`;
  const heading = reminder ? "⏰ Hatırlatma — Onayınızı Bekleyen İzin Talebi" : "📋 Yeni İzin Talebi";
  const introLine = reminder
    ? `<strong>${request.requesterName}</strong> adlı çalışanın izin talebi <strong>6 saattir</strong> onayınızı bekliyor. Lütfen değerlendirin.`
    : `<strong>${request.requesterName}</strong> adlı çalışan size bir izin talebi gönderdi.`;
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
      <h2 style="color:${reminder ? "#c19c00" : "#0078d4"}; margin-bottom: 8px;">${heading}</h2>
      <p>${introLine}</p>
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

// ─── Onaylayıcıya Teams Activity Feed bildirimi ──────────────────────────────

/**
 * Teams'in bildirim merkezine (zil ikonu / activity feed) bildirim gönderir.
 * Permission: TeamsActivity.Send (Application).
 * Manifest'te activityTypes tanımlı olmalı (leaveRequestApproval).
 * Onaylayıcının uygulamayı yüklemiş olması gerekir (Teams Admin Center).
 */
async function notifyApproverInTeams(request, approverId) {
  if (!approverId) return;
  const client = getAppGraphClient();
  const teamsUrl = buildTeamsDeepLink() ||
    process.env.FRONTEND_URL ||
    "https://teams.microsoft.com";

  const previewText = request.leaveType === "saatlik"
    ? `${formatRequestDateRange(request)} saatlik izin onayı bekliyor`
    : `${formatRequestDateRange(request)} izin onayı bekliyor`;

  const payload = {
    topic: {
      source: "text",
      value: "Yeni izin talebi",
      webUrl: teamsUrl,
    },
    activityType: "leaveRequestApproval",
    previewText: { content: previewText },
    templateParameters: [
      { name: "actor", value: request.requesterName || "Bir çalışan" },
    ],
  };

  console.log(`[TEAMS] activity feed bildirimi: user=${approverId}`);

  try {
    await client
      .api(`/users/${approverId}/teamwork/sendActivityNotification`)
      .post(payload);
    console.log(`[TEAMS] BAŞARILI: ${approverId}`);
  } catch (err) {
    console.error(`[TEAMS] BAŞARISIZ:`, {
      message: err.message,
      statusCode: err.statusCode,
      code: err.code,
      body: typeof err.body === "string" ? err.body : JSON.stringify(err.body),
    });
  }
}

// ─── Onaylanan İzin Görüntüleyicilere Bildirim ───────────────────────────────

/**
 * Bir izin onaylandığında, "onaylanan izin görüntüleyici" yetkisine sahip
 * kişilere "X'in izni Y tarafından onaylandı" bildirimi gönderir.
 *
 * @param {object} request   - Onaylanmış talep (status=onaylandi, approverName dolu)
 * @param {Array}  viewers   - [{ id, displayName, mail }] görüntüleyici listesi
 * @param {string} fromEmail - Gönderen kutusu (onaylayanın e-postası)
 */
async function notifyApprovalViewers(request, viewers, fromEmail) {
  if (!Array.isArray(viewers) || viewers.length === 0) return;

  const client = getAppGraphClient();
  const dateLine = formatRequestDateRange(request);
  const sureLine = request.leaveType === "saatlik"
    ? `${request.totalDays} iş günü karşılığı (saatlik)`
    : `${request.totalDays} iş günü`;
  const approverName = request.approverName || "Yöneticisi";
  const subject = `Onaylanan İzin — ${request.requesterName}`;

  const html = `
    <div style="font-family: Segoe UI, Arial, sans-serif; max-width: 560px;">
      <h2 style="color:#107c10; margin-bottom: 8px;">✅ İzin Onaylandı</h2>
      <p><strong>${request.requesterName}</strong> adlı çalışanın izni
         <strong>${approverName}</strong> tarafından onaylandı.</p>
      <table style="border-collapse: collapse; margin: 12px 0;">
        <tr><td style="padding:4px 12px 4px 0;"><b>Çalışan:</b></td><td>${request.requesterName}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>İzin Türü:</b></td><td>${izinTuruLabel(request.leaveType)}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Tarih:</b></td><td>${dateLine}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Süre:</b></td><td>${sureLine}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Onaylayan:</b></td><td>${approverName}</td></tr>
      </table>
      <p style="color:#888; font-size:12px; margin-top:24px;">
        Bu bilgilendirme İzin Onay Sistemi tarafından otomatik gönderildi.
      </p>
    </div>
  `;

  const sender = fromEmail || request.requesterEmail || request.requesterId;

  for (const viewer of viewers) {
    // Kendi onayladığı/açtığı talebi tekrar kendine bildirme
    if (viewer.id === request.approverId || viewer.id === request.requesterId) continue;
    if (!viewer.mail) continue;

    try {
      await client
        .api(`/users/${sender}/sendMail`)
        .post({
          message: {
            subject,
            body: { contentType: "HTML", content: html },
            toRecipients: [{ emailAddress: { address: viewer.mail } }],
          },
          saveToSentItems: false,
        });
      console.log(`[VIEWER MAIL] BAŞARILI: ${viewer.mail}`);
    } catch (err) {
      console.error(`[VIEWER MAIL] BAŞARISIZ to=${viewer.mail}:`, err.message);
    }

    // Teams activity feed (opsiyonel — yetki/manifest varsa)
    try {
      await client
        .api(`/users/${viewer.id}/teamwork/sendActivityNotification`)
        .post({
          topic: {
            source: "text",
            value: "Onaylanan izin",
            webUrl: buildTeamsDeepLink() || process.env.FRONTEND_URL || "https://teams.microsoft.com",
          },
          activityType: "leaveApproved",
          previewText: { content: `${request.requesterName} için izin onaylandı` },
          templateParameters: [
            { name: "actor", value: request.requesterName || "Bir çalışan" },
          ],
        });
      console.log(`[VIEWER TEAMS] BAŞARILI: ${viewer.id}`);
    } catch (err) {
      console.error(`[VIEWER TEAMS] BAŞARISIZ user=${viewer.id}:`, err.message);
    }
  }
}

// ─── Talep Sahibine Sonuç E-postası ──────────────────────────────────────────

/**
 * Talep işlendiğinde (onay/red) talebi açan çalışana e-posta gönderir.
 * Teams DM'e ek olarak — kullanıcı maili de görsün.
 *
 * @param {object} request   - Güncellenmiş talep (status onaylandi/reddedildi)
 * @param {string} fromEmail - Gönderen kutusu (onaylayanın e-postası)
 */
async function notifyRequesterByEmail(request, fromEmail) {
  const to = request.requesterEmail;
  if (!to) {
    console.warn("notifyRequesterByEmail: talep sahibinin maili yok, atlandı.");
    return;
  }
  const client = getAppGraphClient();
  const isApproved = request.status === "onaylandi";
  const dateLine = formatRequestDateRange(request);
  const sureLine = request.leaveType === "saatlik"
    ? `${request.totalDays} iş günü karşılığı (saatlik)`
    : `${request.totalDays} iş günü`;
  const approverName = request.approverName || "Yöneticiniz";
  const subject = isApproved
    ? "✅ İzin Talebiniz Onaylandı"
    : "❌ İzin Talebiniz Reddedildi";

  const html = `
    <div style="font-family: Segoe UI, Arial, sans-serif; max-width: 560px;">
      <h2 style="color:${isApproved ? "#107c10" : "#a4262c"}; margin-bottom: 8px;">${subject}</h2>
      <p><strong>${request.requesterName}</strong>, izin talebiniz
         <strong>${approverName}</strong> tarafından ${isApproved ? "onaylandı" : "reddedildi"}.</p>
      <table style="border-collapse: collapse; margin: 12px 0;">
        <tr><td style="padding:4px 12px 4px 0;"><b>İzin Türü:</b></td><td>${izinTuruLabel(request.leaveType)}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Tarih:</b></td><td>${dateLine}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>Süre:</b></td><td>${sureLine}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;"><b>İşlem Yapan:</b></td><td>${approverName}</td></tr>
        ${request.approverComment ? `<tr><td style="padding:4px 12px 4px 0;vertical-align:top;"><b>Yorum:</b></td><td>${request.approverComment}</td></tr>` : ""}
      </table>
      <p style="color:#888; font-size:12px; margin-top:24px;">
        Bu e-posta İzin Onay Sistemi tarafından otomatik gönderildi.
      </p>
    </div>
  `;

  // Onaylayanın kutusundan gönder; yoksa talep sahibinin kendi kutusuna düş
  const senderUserId = fromEmail || to;

  console.log(`[REQUESTER MAIL] gönderiliyor: from=${senderUserId} to=${to} status=${request.status}`);

  try {
    await client
      .api(`/users/${senderUserId}/sendMail`)
      .post({
        message: {
          subject,
          body: { contentType: "HTML", content: html },
          toRecipients: [{ emailAddress: { address: to } }],
        },
        saveToSentItems: false,
      });
    console.log(`[REQUESTER MAIL] BAŞARILI: ${to}`);
  } catch (err) {
    console.error(`[REQUESTER MAIL] BAŞARISIZ from=${senderUserId} to=${to}:`, err.message);
  }
}

module.exports = { notifyChannel, notifyUser, notifyApproverByEmail, notifyApproverInTeams, notifyApprovalViewers, notifyRequesterByEmail };
