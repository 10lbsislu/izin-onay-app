/**
 * requestService.ts
 * Azure Functions API ile iletişim kurar.
 * Tüm isteklere Teams SSO Bearer token eklenir.
 */

import type { LeaveRequest, Approver, OrgUser } from "../types";

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
  payload: Omit<LeaveRequest, "id" | "status" | "createdAt" | "updatedAt">
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

// ─── Onaylayıcılar ────────────────────────────────────────────────────────────

export async function getApprovers(token: string): Promise<Approver[]> {
  const res = await apiFetch<{ approvers: Approver[] }>(
    "/getApprovers",
    { method: "GET" },
    token
  );
  return res.approvers;
}

export async function addApprover(token: string, user: OrgUser): Promise<Approver> {
  const res = await apiFetch<{ approver: Approver }>(
    "/manageApprovers",
    {
      method: "POST",
      body: JSON.stringify({
        action: "add",
        userId: user.id,
        displayName: user.displayName,
        mail: user.mail,
      }),
    },
    token
  );
  return res.approver;
}

export async function removeApprover(token: string, approverId: string): Promise<void> {
  await apiFetch(
    "/manageApprovers",
    { method: "POST", body: JSON.stringify({ action: "remove", userId: approverId }) },
    token
  );
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
