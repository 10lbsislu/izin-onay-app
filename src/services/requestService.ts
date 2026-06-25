/**
 * requestService.ts
 * Azure Functions API ile iletişim kurar.
 * Tüm isteklere Teams SSO Bearer token eklenir.
 */

import type { LeaveRequest, Approver, OrgUser, HierarchyNode, TreeNode } from "../types";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "/api";

// ─── Görünürlük meta tipi ─────────────────────────────────────────────────────
export interface VisibilityMeta {
  badge: "only_you_and_approvers" | "approver_only" | "normal";
  label: string | null;
  icon: "lock" | null;
}

export interface EnrichedLeaveRequest extends LeaveRequest {
  _visibility?: VisibilityMeta;
}

export interface GetRequestsResponse {
  requests: EnrichedLeaveRequest[];
  meta: {
    callerId: string;
    isApprover: boolean;
    total: number;
  };
}

// ─── Authenticated fetch ──────────────────────────────────────────────────────
async function apiFetch<T>(
  endpoint: string,
  options: RequestInit = {},
  token?: string
): Promise<T> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    ...(token ? { Authorization: `Bearer ${token}` } : {}),
    ...(options.headers as Record<string, string>),
  };

  const res = await fetch(`${API_BASE}${endpoint}`, { ...options, headers });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(err.error || `HTTP ${res.status}`);
  }

  return res.json();
}

// ─── İzin Talepleri ──────────────────────────────────────────────────────────

/** Görünürlük kuralları sunucu tarafında uygulanır — artık userId parametresi yok */
export async function getRequests(token: string): Promise<GetRequestsResponse> {
  return apiFetch<GetRequestsResponse>("/getRequests", { method: "GET" }, token);
}

export async function createRequest(
  token: string,
  payload: Omit<LeaveRequest, "id" | "status" | "createdAt" | "updatedAt"> & { approverId?: string }
): Promise<LeaveRequest> {
  const res = await apiFetch<{ request: LeaveRequest }>(
    "/createRequest",
    { method: "POST", body: JSON.stringify(payload) },
    token
  );
  return res.request;
}

export async function updateRequestStatus(
  token: string,
  requestId: string,
  status: "onaylandi" | "reddedildi",
  approverComment: string
): Promise<LeaveRequest> {
  // approverId/approverName artık backend'de token'dan alınıyor
  const res = await apiFetch<{ request: LeaveRequest }>(
    "/updateRequest",
    {
      method: "PUT",
      body: JSON.stringify({ id: requestId, status, approverComment }),
    },
    token
  );
  return res.request;
}

// ─── Onaylanan İzinler (yetkili görüntüleyiciler) ─────────────────────────────

export interface ApprovedLeavesResponse {
  requests: LeaveRequest[];
  meta: {
    callerId: string;
    total: number;
  };
}

/** Tüm onaylanan izinleri döndürür — sadece isApprovalViewer / tree admin çağırabilir */
export async function getApprovedLeaves(token: string): Promise<LeaveRequest[]> {
  const res = await apiFetch<ApprovedLeavesResponse>(
    "/getApprovedLeaves",
    { method: "GET" },
    token
  );
  return res.requests;
}

// ─── Takvim (izinler + doğum günleri) ─────────────────────────────────────────

export interface CalendarLeave {
  id: string;
  requesterName: string;
  leaveType: LeaveRequest["leaveType"];
  startDate: string;
  endDate: string;
  startTime?: string;
  endTime?: string;
}

export interface Birthday {
  id: string;
  name: string;
  birthDate: string; // "MM-DD"
}

export interface CalendarResponse {
  leaves: CalendarLeave[];
  birthdays: Birthday[];
}

/** Takvim verisi — herkese açık (kimliği doğrulanmış her kullanıcı) */
export async function getCalendar(token: string): Promise<CalendarResponse> {
  return apiFetch<CalendarResponse>("/getCalendar", { method: "GET" }, token);
}

export async function addBirthday(token: string, name: string, birthDate: string): Promise<Birthday> {
  const res = await apiFetch<{ birthday: Birthday }>(
    "/manageBirthdays",
    { method: "POST", body: JSON.stringify({ action: "add", name, birthDate }) },
    token
  );
  return res.birthday;
}

export async function removeBirthday(token: string, id: string): Promise<void> {
  await apiFetch(
    "/manageBirthdays",
    { method: "POST", body: JSON.stringify({ action: "remove", id }) },
    token
  );
}

// ─── Hiyerarşi ────────────────────────────────────────────────────────────────

export async function getHierarchy(token: string): Promise<HierarchyNode[]> {
  const res = await apiFetch<{ nodes: HierarchyNode[] }>(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "get-tree" }) },
    token
  );
  return res.nodes;
}

export async function setUserManager(
  token: string,
  user: OrgUser,
  managerId: string
): Promise<HierarchyNode> {
  const res = await apiFetch<{ node: HierarchyNode }>(
    "/manageApprovers",
    {
      method: "POST",
      body: JSON.stringify({
        action: "set-manager",
        userId: user.id,
        managerId,
        displayName: user.displayName,
        mail: user.mail,
      }),
    },
    token
  );
  return res.node;
}

export async function removeHierarchyNode(token: string, userId: string): Promise<void> {
  await apiFetch(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "remove", userId }) },
    token
  );
}

export async function grantApprovalViewer(token: string, userId: string): Promise<HierarchyNode> {
  const res = await apiFetch<{ node: HierarchyNode }>(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "grant-viewer", userId }) },
    token
  );
  return res.node;
}

export async function revokeApprovalViewer(token: string, userId: string): Promise<HierarchyNode> {
  const res = await apiFetch<{ node: HierarchyNode }>(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "revoke-viewer", userId }) },
    token
  );
  return res.node;
}

export async function grantTreeAdmin(token: string, userId: string): Promise<HierarchyNode> {
  const res = await apiFetch<{ node: HierarchyNode }>(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "grant-admin", userId }) },
    token
  );
  return res.node;
}

export async function revokeTreeAdmin(token: string, userId: string): Promise<HierarchyNode> {
  const res = await apiFetch<{ node: HierarchyNode }>(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "revoke-admin", userId }) },
    token
  );
  return res.node;
}

// Tree builder helper
export function buildTree(nodes: HierarchyNode[]): TreeNode[] {
  const byId = new Map<string, TreeNode>();
  nodes.forEach((n) => byId.set(n.id, { ...n, children: [] }));
  const roots: TreeNode[] = [];
  byId.forEach((node) => {
    if (node.managerId && byId.has(node.managerId)) {
      byId.get(node.managerId)!.children.push(node);
    } else {
      roots.push(node);
    }
  });
  // İsme göre sırala (her seviyede)
  const sortRec = (arr: TreeNode[]) => {
    arr.sort((a, b) => a.displayName.localeCompare(b.displayName, "tr"));
    arr.forEach((n) => sortRec(n.children));
  };
  sortRec(roots);
  return roots;
}

// Bir node'un altındaki tüm id'ler (cycle prevention için)
export function getDescendantIds(nodes: HierarchyNode[], rootId: string): Set<string> {
  const descendants = new Set<string>();
  const recurse = (id: string) => {
    nodes.filter((n) => n.managerId === id).forEach((child) => {
      if (!descendants.has(child.id)) {
        descendants.add(child.id);
        recurse(child.id);
      }
    });
  };
  recurse(rootId);
  return descendants;
}

// Eski API uyumluluğu — yeni kod kullanmamalı
export async function getApprovers(token: string): Promise<Approver[]> {
  const res = await apiFetch<{ approvers: Approver[] }>(
    "/getApprovers",
    { method: "GET" },
    token
  );
  return res.approvers;
}

// ─── Kullanıcı Arama ─────────────────────────────────────────────────────────

export async function searchOrgUsers(token: string, query: string): Promise<OrgUser[]> {
  if (query.trim().length < 2) return [];
  const res = await apiFetch<{ users: OrgUser[] }>(
    `/searchUsers?q=${encodeURIComponent(query)}`,
    { method: "GET" },
    token
  );
  return res.users;
}
