// ─── Kullanıcı ───────────────────────────────────────────────────────────────
export interface OrgUser {
  id: string;
  displayName: string;
  mail: string;
  jobTitle?: string;
  department?: string;
  userPrincipalName?: string;
}

// ─── Onaylayıcı ───────────────────────────────────────────────────────────────
export interface Approver {
  id: string;
  displayName: string;
  mail: string;
  addedAt: string;
}

// ─── İzin Talebi ─────────────────────────────────────────────────────────────
export type LeaveType =
  | "yillik"
  | "hastalik"
  | "mazeret"
  | "ucretsiz"
  | "diger";

export type LeaveStatus = "beklemede" | "onaylandi" | "reddedildi";

export interface LeaveRequest {
  id: string;
  requesterId: string;
  requesterName: string;
  requesterEmail: string;
  leaveType: LeaveType;
  startDate: string;       // ISO date string YYYY-MM-DD
  endDate: string;         // ISO date string YYYY-MM-DD
  totalDays: number;
  description: string;
  status: LeaveStatus;
  approverId?: string;
  approverName?: string;
  approverComment?: string;
  createdAt: string;       // ISO datetime
  updatedAt?: string;      // ISO datetime
}

// ─── API Yanıtları ────────────────────────────────────────────────────────────
export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
}

// ─── Uygulama Durumu ──────────────────────────────────────────────────────────
export type AppView = "employee" | "admin";

export interface AppContextType {
  currentUser: OrgUser | null;
  isAdmin: boolean;
  teamsToken: string | null;
}

// ─── İzin Türü Etiketi ────────────────────────────────────────────────────────
export const LEAVE_TYPE_LABELS: Record<LeaveType, string> = {
  yillik: "Yıllık İzin",
  hastalik: "Hastalık İzni",
  mazeret: "Mazeret İzni",
  ucretsiz: "Ücretsiz İzin",
  diger: "Diğer",
};

export const STATUS_LABELS: Record<LeaveStatus, string> = {
  beklemede: "Beklemede",
  onaylandi: "Onaylandı",
  reddedildi: "Reddedildi",
};
