/**
 * shared/excelService.js
 * SharePoint Excel dosyasına Graph Workbook API ile erişim.
 *
 * Excel yapısı:
 *   Sheet "Talepler"    → Tablo: TaleplerTablosu
 *   Sheet "Onaylayıcılar" → Tablo: OnaylayicilarTablosu
 *
 * Kurulum notları:
 *   - SharePoint'te bir Excel dosyası oluşturun
 *   - Her iki sayfada da adlandırılmış tablo oluşturun
 *   - Dosya ID'sini env değişkenine yazın
 */

const { getAppGraphClient } = require("./graphClient");

const SITE_ID = process.env.SHAREPOINT_SITE_ID;
const FILE_ID = process.env.EXCEL_FILE_ID;

// ─── Workbook URL yardımcısı ─────────────────────────────────────────────────
const workbookBase = () =>
  `/sites/${SITE_ID}/drive/items/${FILE_ID}/workbook`;

// ─── Sütun eşlemeleri ────────────────────────────────────────────────────────
const REQUEST_COLS = [
  "id", "requesterId", "requesterName", "requesterEmail",
  "leaveType", "startDate", "endDate", "totalDays",
  "description", "status", "approverId", "approverName",
  "approverComment", "createdAt", "updatedAt",
];

const APPROVER_COLS = ["id", "displayName", "mail", "addedAt"];

function excelDateToISO(value) {
  if (!value) return "";
  if (typeof value === "string" && value.includes("-")) return value;
  const num = Number(value);
  if (isNaN(num)) return value;
  // Excel epoch: 1 Ocak 1900 = 1, ancak Excel 1900'ü yanlış leap year sayar
  const date = new Date((num - 25569) * 86400 * 1000);
  return date.toISOString().split("T")[0];
}

function rowToRequest(row) {
  const obj = {};
  REQUEST_COLS.forEach((col, i) => {
    obj[col] = row[i] ?? "";
  });
  if (obj.totalDays) obj.totalDays = Number(obj.totalDays);
  if (obj.startDate) obj.startDate = excelDateToISO(obj.startDate);
  if (obj.endDate) obj.endDate = excelDateToISO(obj.endDate);
  return obj;
}

function requestToRow(req) {
  return REQUEST_COLS.map((col) => {
    const v = req[col];
    return v === undefined || v === null ? "" : String(v);
  });
}

function rowToApprover(row) {
  const obj = {};
  APPROVER_COLS.forEach((col, i) => {
    obj[col] = row[i] ?? "";
  });
  return obj;
}

// ─── İzin Talepleri ──────────────────────────────────────────────────────────

/** Tüm talepleri getir */
async function getAllRequests() {
  const client = getAppGraphClient();
  try {
    const res = await client
      .api(`${workbookBase()}/tables/TaleplerTablosu/rows`)
      .get();
    return (res.value || []).map((r) => rowToRequest(r.values[0]));
  } catch (err) {
    console.error("Excel getRequests error:", err);
    throw err;
  }
}

/** Belirli kullanıcının taleplerini getir */
async function getRequestsByUser(userId) {
  const all = await getAllRequests();
  return all.filter((r) => r.requesterId === userId);
}

/** Yeni talep satırı ekle */
async function addRequest(request) {
  const client = getAppGraphClient();
  await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows/add`)
    .post({ values: [requestToRow(request)] });
  return request;
}

/** Mevcut talebi güncelle (satır bul → patch) */
async function updateRequest(requestId, updates) {
  const client = getAppGraphClient();

  // Tüm satırları getir, ID'ye göre bul
  const res = await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows`)
    .get();

  const rows = res.value || [];
  const rowIndex = rows.findIndex((r) => r.values[0][0] === requestId);

  if (rowIndex === -1) throw new Error(`Talep bulunamadı: ${requestId}`);

  // Mevcut veriyi güncelle
  const current = rowToRequest(rows[rowIndex].values[0]);
  const updated = { ...current, ...updates };

  await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows/itemAt(index=${rowIndex})`)
    .patch({ values: [requestToRow(updated)] });

  return updated;
}

// ─── Onaylayıcılar ────────────────────────────────────────────────────────────

/** Onaylayıcı listesini getir */
async function getAllApprovers() {
  const client = getAppGraphClient();
  try {
    const SITE_ID = process.env.SHAREPOINT_SITE_ID;
    const FILE_ID = process.env.EXCEL_FILE_ID;
    const res = await client
      .api(`/sites/${SITE_ID}/drive/items/${FILE_ID}/workbook/tables/OnaylayicilarTablosu/rows`)
      .get();
    return (res.value || []).map((r) => rowToApprover(r.values[0]));
  } catch (err) {
    console.error("Excel getAllApprovers error:", err.message);
    return [];
  }
}

/** Onaylayıcı ekle */
async function addApprover(approver) {
  const client = getAppGraphClient();
  const row = APPROVER_COLS.map((col) => approver[col] ?? "");
  await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/add`)
    .post({ values: [row] });
  return approver;
}

/** Onaylayıcı kaldır */
async function removeApprover(approverId) {
  const client = getAppGraphClient();

  const res = await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows`)
    .get();

  const rows = res.value || [];
  const rowIndex = rows.findIndex((r) => r.values[0][0] === approverId);

  if (rowIndex === -1) throw new Error(`Onaylayıcı bulunamadı: ${approverId}`);

  await client
    .api(
      `${workbookBase()}/tables/OnaylayicilarTablosu/rows/itemAt(index=${rowIndex})`
    )
    .delete();
}

module.exports = {
  getAllRequests,
  getRequestsByUser,
  addRequest,
  updateRequest,
  getAllApprovers,
  addApprover,
  removeApprover,
};
