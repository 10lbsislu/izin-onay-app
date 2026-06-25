/**
 * shared/excelService.js
 * SharePoint Excel dosyasına Graph Workbook API ile erişim.
 *
 * Excel yapısı:
 *   Sheet "Talepler"      → Tablo: TaleplerTablosu
 *   Sheet "Onaylayicilar" → Tablo: OnaylayicilarTablosu (hiyerarşi düğümleri)
 */

const { getAppGraphClient } = require("./graphClient");

const SITE_ID = process.env.SHAREPOINT_SITE_ID;
const FILE_ID = process.env.EXCEL_FILE_ID;

const workbookBase = () =>
  `/sites/${SITE_ID}/drive/items/${FILE_ID}/workbook`;

const REQUEST_COLS = [
  "id", "requesterId", "requesterName", "requesterEmail",
  "leaveType", "startDate", "endDate", "totalDays",
  "description", "status", "approverId", "approverName",
  "approverComment", "createdAt", "updatedAt",
  "startTime", "endTime",
];

const APPROVER_COLS = ["id", "displayName", "mail", "managerId", "isTreeAdmin", "addedAt", "isApprovalViewer"];

function excelDateToISO(value) {
  if (!value) return "";
  if (typeof value === "string" && value.includes("-")) return value;
  const num = Number(value);
  if (isNaN(num)) return value;
  const date = new Date((num - 25569) * 86400 * 1000);
  return date.toISOString().split("T")[0];
}

// "15:00" stringi olarak geldi → olduğu gibi döndür
// 0.625 sayısı (Excel time serial) olarak geldi → "15:00"e çevir
function excelTimeToHHMM(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "string" && value.includes(":")) return value;
  const num = Number(value);
  if (isNaN(num)) return String(value);
  const totalMinutes = Math.round(num * 24 * 60);
  const hours = Math.floor(totalMinutes / 60) % 24;
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function rowToRequest(row) {
  const obj = {};
  REQUEST_COLS.forEach((col, i) => {
    obj[col] = row[i] ?? "";
  });
  if (obj.totalDays) obj.totalDays = Number(obj.totalDays);
  if (obj.startDate) obj.startDate = excelDateToISO(obj.startDate);
  if (obj.endDate) obj.endDate = excelDateToISO(obj.endDate);
  if (obj.startTime !== "" && obj.startTime !== undefined) obj.startTime = excelTimeToHHMM(obj.startTime);
  if (obj.endTime !== "" && obj.endTime !== undefined) obj.endTime = excelTimeToHHMM(obj.endTime);
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

async function getRequestsByUser(userId) {
  const all = await getAllRequests();
  return all.filter((r) => r.requesterId === userId);
}

async function addRequest(request) {
  const client = getAppGraphClient();
  await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows/add`)
    .post({ values: [requestToRow(request)] });
  return request;
}

async function updateRequest(requestId, updates) {
  const client = getAppGraphClient();

  const res = await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows`)
    .get();

  const rows = res.value || [];
  const rowIndex = rows.findIndex((r) => r.values[0][0] === requestId);

  if (rowIndex === -1) throw new Error(`Talep bulunamadı: ${requestId}`);

  const current = rowToRequest(rows[rowIndex].values[0]);
  const updated = { ...current, ...updates };

  await client
    .api(`${workbookBase()}/tables/TaleplerTablosu/rows/itemAt(index=${rowIndex})`)
    .patch({ values: [requestToRow(updated)] });

  return updated;
}

// ─── Hiyerarşi (eski "Onaylayicilar" tablosu, yeni semantik) ─────────────────

async function getAllApprovers() {
  const client = getAppGraphClient();
  try {
    const res = await client
      .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows`)
      .get();
    return (res.value || []).map((r) => rowToApprover(r.values[0]));
  } catch (err) {
    console.error("Excel getAllApprovers error:", err.message);
    return [];
  }
}

async function getHierarchyNode(userId) {
  const all = await getAllApprovers();
  return all.find((n) => n.id === userId) || null;
}

async function getDirectReports(managerId) {
  const all = await getAllApprovers();
  return all.filter((n) => n.managerId === managerId);
}

async function getManagerOf(userId) {
  const node = await getHierarchyNode(userId);
  return node ? node.managerId || "" : "";
}

async function isTreeAdmin(userId) {
  const node = await getHierarchyNode(userId);
  return !!node && String(node.isTreeAdmin).toLowerCase() === "true";
}

// newManagerId zincirinde userId'ye ulaşılıyorsa döngü var
async function detectCycle(userId, newManagerId) {
  if (!newManagerId) return false;
  if (newManagerId === userId) return true;
  const all = await getAllApprovers();
  const byId = new Map(all.map((n) => [n.id, n]));
  let cursor = newManagerId;
  const seen = new Set();
  while (cursor) {
    if (cursor === userId) return true;
    if (seen.has(cursor)) return true;
    seen.add(cursor);
    const node = byId.get(cursor);
    cursor = node ? node.managerId : "";
  }
  return false;
}

// Mevcut satırın index'ini bul (yoksa -1)
async function _findRowIndex(userId) {
  const client = getAppGraphClient();
  const res = await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows`)
    .get();
  const rows = res.value || [];
  return { rows, index: rows.findIndex((r) => r.values[0][0] === userId) };
}

async function setUserManager(userId, managerId, displayName, mail) {
  if (await detectCycle(userId, managerId)) {
    throw new Error("Döngü: bir kullanıcı kendi astının altına alınamaz.");
  }
  const client = getAppGraphClient();
  const { rows, index } = await _findRowIndex(userId);

  if (index === -1) {
    const newRow = {
      id: userId,
      displayName: displayName || "",
      mail: mail || "",
      managerId: managerId || "",
      isTreeAdmin: "",
      addedAt: new Date().toISOString(),
    };
    const row = APPROVER_COLS.map((col) => newRow[col] ?? "");
    await client
      .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/add`)
      .post({ values: [row] });
    return newRow;
  }

  const current = rowToApprover(rows[index].values[0]);
  const updated = {
    ...current,
    managerId: managerId || "",
    displayName: displayName || current.displayName,
    mail: mail || current.mail,
  };
  const row = APPROVER_COLS.map((col) => updated[col] ?? "");
  await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/itemAt(index=${index})`)
    .patch({ values: [row] });
  return updated;
}

async function setTreeAdmin(userId, value) {
  const client = getAppGraphClient();
  const { rows, index } = await _findRowIndex(userId);
  if (index === -1) throw new Error(`Kullanıcı hiyerarşide yok: ${userId}`);
  const current = rowToApprover(rows[index].values[0]);
  const updated = { ...current, isTreeAdmin: value ? "true" : "" };
  const row = APPROVER_COLS.map((col) => updated[col] ?? "");
  await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/itemAt(index=${index})`)
    .patch({ values: [row] });
  return updated;
}

async function removeHierarchyNode(userId) {
  const children = await getDirectReports(userId);
  if (children.length > 0) {
    throw new Error(`Bu kullanıcının ${children.length} alt çalışanı var. Önce onları başka bir amire taşıyın.`);
  }
  const client = getAppGraphClient();
  const { index } = await _findRowIndex(userId);
  if (index === -1) throw new Error(`Kullanıcı bulunamadı: ${userId}`);
  await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/itemAt(index=${index})`)
    .delete();
}

async function countTreeAdmins() {
  const all = await getAllApprovers();
  return all.filter((n) => String(n.isTreeAdmin).toLowerCase() === "true").length;
}

// ─── Onaylanan İzin Görüntüleyiciler ─────────────────────────────────────────
// isApprovalViewer = "true" olan hiyerarşi düğümleri tüm onaylanan izinleri görebilir.

async function getApprovalViewers() {
  const all = await getAllApprovers();
  return all.filter((n) => String(n.isApprovalViewer).toLowerCase() === "true");
}

async function isApprovalViewer(userId) {
  const node = await getHierarchyNode(userId);
  return !!node && String(node.isApprovalViewer).toLowerCase() === "true";
}

async function setApprovalViewer(userId, value) {
  const client = getAppGraphClient();
  const { rows, index } = await _findRowIndex(userId);
  if (index === -1) throw new Error(`Kullanıcı hiyerarşide yok: ${userId}`);
  const current = rowToApprover(rows[index].values[0]);
  const updated = { ...current, isApprovalViewer: value ? "true" : "" };
  const row = APPROVER_COLS.map((col) => updated[col] ?? "");
  await client
    .api(`${workbookBase()}/tables/OnaylayicilarTablosu/rows/itemAt(index=${index})`)
    .patch({ values: [row] });
  return updated;
}

// ─── Doğum Günleri ───────────────────────────────────────────────────────────
// Sheet "DogumGunleri" → Tablo: DogumGunleriTablosu
// Kolonlar: id | name | birthDate ("MM-DD" formatında, yıl önemsiz)

const BIRTHDAY_COLS = ["id", "name", "birthDate"];

function rowToBirthday(row) {
  const obj = {};
  BIRTHDAY_COLS.forEach((col, i) => { obj[col] = row[i] ?? ""; });
  return obj;
}

async function getAllBirthdays() {
  const client = getAppGraphClient();
  try {
    const res = await client
      .api(`${workbookBase()}/tables/DogumGunleriTablosu/rows`)
      .get();
    return (res.value || []).map((r) => rowToBirthday(r.values[0]));
  } catch (err) {
    console.error("getAllBirthdays error:", err.message);
    return [];
  }
}

async function addBirthday(birthday) {
  const client = getAppGraphClient();
  const row = BIRTHDAY_COLS.map((col) => birthday[col] ?? "");
  await client
    .api(`${workbookBase()}/tables/DogumGunleriTablosu/rows/add`)
    .post({ values: [row] });
  return birthday;
}

async function removeBirthday(id) {
  const client = getAppGraphClient();
  const res = await client.api(`${workbookBase()}/tables/DogumGunleriTablosu/rows`).get();
  const rows = res.value || [];
  const index = rows.findIndex((r) => r.values[0][0] === id);
  if (index === -1) throw new Error(`Doğum günü kaydı bulunamadı: ${id}`);
  await client
    .api(`${workbookBase()}/tables/DogumGunleriTablosu/rows/itemAt(index=${index})`)
    .delete();
}

module.exports = {
  getAllRequests,
  getRequestsByUser,
  addRequest,
  updateRequest,
  getAllApprovers,
  getHierarchyNode,
  getDirectReports,
  getManagerOf,
  isTreeAdmin,
  detectCycle,
  setUserManager,
  setTreeAdmin,
  removeHierarchyNode,
  countTreeAdmins,
  getApprovalViewers,
  isApprovalViewer,
  setApprovalViewer,
  getAllBirthdays,
  addBirthday,
  removeBirthday,
};
