// ─── Kullanıcı ───────────────────────────────────────────────────────────────
export interface OrgUser {
  id: string;
  displayName: string;
  mail: string;
  jobTitle?: string;
  department?: string;
  userPrincipalName?: string;
}

// ─── Hiyerarşi Düğümü ─────────────────────────────────────────────────────────
export interface HierarchyNode {
  id: string;
  displayName: string;
  mail: string;
  managerId: string;     // "" = root
  isTreeAdmin: boolean;
  addedAt: string;
}

export interface TreeNode extends HierarchyNode {
  children: TreeNode[];
}

// Geriye uyum için alias
export type Approver = HierarchyNode;

// ─── İzin Talebi ─────────────────────────────────────────────────────────────
export type LeaveType =
  | "yillik"
  | "saatlik"
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
  startTime?: string;      // HH:MM (yalnız saatlik izinde)
  endTime?: string;        // HH:MM (yalnız saatlik izinde)
  totalDays: number;       // saatlik için kesirli olabilir (saat/8)
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
  isTreeAdmin: boolean;
  hasDirectReports: boolean;
  isRoot: boolean;
  teamsToken: string | null;
}

// ─── İzin Türü Etiketi ────────────────────────────────────────────────────────
export const LEAVE_TYPE_LABELS: Record<LeaveType, string> = {
  yillik: "Yıllık İzin",
  saatlik: "Saatlik İzin",
  ucretsiz: "Ücretsiz İzin",
  diger: "Diğer",
};

export const STATUS_LABELS: Record<LeaveStatus, string> = {
  beklemede: "Beklemede",
  onaylandi: "Onaylandı",
  reddedildi: "Reddedildi",
};
