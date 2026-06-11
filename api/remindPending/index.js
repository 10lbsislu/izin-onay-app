/**
 * remindPending/index.js
 * Timer Trigger — her saat başı çalışır.
 *
 * 6 saattir "beklemede" olan (onaylanmamış) talepler için onaylayıcıya
 * BİR KEZ hatırlatma e-postası gönderir.
 *
 * "Tek sefer" mantığı: ekstra bir Excel kolonu tutmadan, talebin yaşı
 * [6 saat, 7 saat) penceresine girdiği TEK saatlik çalışmada mail atılır.
 * Timer saatlik olduğu için her talep bu pencereye yalnızca bir kez düşer.
 */

const { app } = require("@azure/functions");
const { getAllRequests, getHierarchyNode } = require("../shared/excelService");
const { notifyApproverByEmail } = require("../shared/notifyService");

const ONE_HOUR = 60 * 60 * 1000;
const SIX_HOURS = 6 * ONE_HOUR;

app.timer("remindPending", {
  schedule: "0 0 * * * *", // her saat başı (sn dk saat gün ay haftagünü)
  handler: async (myTimer, context) => {
    const now = Date.now();
    context.log("[REMIND] bekleyen talepler taranıyor...");

    try {
      const all = await getAllRequests();
      const pending = all.filter((r) => r.status === "beklemede" && r.createdAt);

      let sent = 0;
      for (const req of pending) {
        const age = now - new Date(req.createdAt).getTime();

        // 6. saati yeni geçenler: [6h, 7h) — tek seferlik
        if (age < SIX_HOURS || age >= SIX_HOURS + ONE_HOUR) continue;

        if (!req.approverId) {
          context.log(`[REMIND] atlandı (atanmış amir yok): ${req.id}`);
          continue;
        }

        const approverNode = await getHierarchyNode(req.approverId);
        const approverEmail = approverNode ? approverNode.mail : "";
        if (!approverEmail) {
          context.log(`[REMIND] onaylayıcı maili bulunamadı: ${req.id}`);
          continue;
        }

        try {
          await notifyApproverByEmail(req, approverEmail, { reminder: true });
          sent++;
          context.log(`[REMIND] hatırlatma gönderildi: ${req.id} -> ${approverEmail}`);
        } catch (err) {
          context.error(`[REMIND] mail hatası ${req.id}:`, err.message);
        }
      }

      context.log(`[REMIND] tamamlandı. ${pending.length} bekleyen talep, ${sent} hatırlatma gönderildi.`);
    } catch (err) {
      context.error("[REMIND] genel hata:", err);
    }
  },
});
