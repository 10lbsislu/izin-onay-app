/**
 * AdminPanel.tsx
 * Yönetici paneli.
 * updateRequestStatus artık approverId/approverName almıyor — token'dan gelir.
 */

import React, { useEffect, useState, useCallback } from "react";
import {
  TabList, Tab, Text, Spinner, Button, Card, Badge, Avatar,
  Dialog, DialogTrigger, DialogSurface, DialogTitle, DialogBody,
  DialogActions, Field, Textarea, Select,
  Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell,
  Toolbar, ToolbarButton, makeStyles, tokens,
  Toast, ToastTitle, ToastBody, useToastController, useId, Toaster,
} from "@fluentui/react-components";
import {
  CheckmarkCircleRegular, DismissCircleRegular,
  ArrowClockwiseRegular, PeopleRegular,
  ClipboardTaskListLtrRegular, LockClosedRegular,
} from "@fluentui/react-icons";
import type { OrgUser, LeaveStatus } from "../types";
import type { EnrichedLeaveRequest, GetRequestsResponse } from "../services/requestService";
import { LEAVE_TYPE_LABELS } from "../types";
import { getRequests, updateRequestStatus } from "../services/requestService";
import { StatusBadge } from "./StatusBadge";
import { VisibilityBadge } from "./VisibilityBadge";
import { HierarchyEditor } from "./HierarchyEditor";

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", gap: "16px" },
  requestCard: { border: `1px solid ${tokens.colorNeutralStroke1}`, marginBottom: "8px" },
  pendingCard: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    marginBottom: "8px",
    borderLeft: `3px solid ${tokens.colorPaletteYellowBorderActive}`,
  },
  cardBody: { display: "flex", flexDirection: "column", gap: "8px", padding: "4px 0" },
  metaRow: { display: "flex", flexWrap: "wrap", gap: "10px", alignItems: "center" },
  actions: { display: "flex", gap: "8px", flexWrap: "wrap", marginTop: "4px" },
  filterBar: { display: "flex", gap: "12px", alignItems: "center", flexWrap: "wrap" },
  emptyState: {
    display: "flex", flexDirection: "column", alignItems: "center",
    gap: "12px", padding: "48px 24px", color: tokens.colorNeutralForeground3,
  },
  visibilityNote: {
    display: "flex", alignItems: "center", gap: "6px",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
  },
});

type AdminTab = "pending" | "all" | "hierarchy";

interface AdminPanelProps {
  currentUser: OrgUser;
  token: string;
  isTreeAdmin: boolean;
  hasDirectReports: boolean;
}

export const AdminPanel: React.FC<AdminPanelProps> = ({ currentUser, token, isTreeAdmin, hasDirectReports }) => {
  const styles = useStyles();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const defaultTab: AdminTab = hasDirectReports ? "pending" : (isTreeAdmin ? "hierarchy" : "all");
  const [activeTab, setActiveTab] = useState<AdminTab>(defaultTab);
  const [requests, setRequests] = useState<EnrichedLeaveRequest[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [filterStatus, setFilterStatus] = useState<LeaveStatus | "hepsi">("hepsi");

  const [selectedRequest, setSelectedRequest] = useState<EnrichedLeaveRequest | null>(null);
  const [dialogAction, setDialogAction] = useState<"approve" | "reject" | null>(null);
  const [approverComment, setApproverComment] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);

  const loadData = useCallback(async () => {
    setIsLoading(true);
    try {
      const res: GetRequestsResponse = await getRequests(token);
      setRequests(res.requests.sort((a, b) => b.createdAt.localeCompare(a.createdAt)));
    } catch (err) {
      dispatchToast(
        <Toast><ToastTitle>Yükleme Hatası</ToastTitle><ToastBody>{String(err)}</ToastBody></Toast>,
        { intent: "error" }
      );
    } finally {
      setIsLoading(false);
    }
  }, [token, dispatchToast]);

  useEffect(() => { loadData(); }, [loadData]);

  const pendingRequests = requests.filter((r) => r.status === "beklemede");
  const filteredRequests = filterStatus === "hepsi"
    ? requests
    : requests.filter((r) => r.status === filterStatus);

  // ── Onayla / Reddet ──────────────────────────────────────────────────────
  const handleAction = async () => {
    if (!selectedRequest || !dialogAction) return;
    setIsProcessing(true);
    try {
      // approverId/approverName artık backend'de token'dan alınır
      await updateRequestStatus(
        token,
        selectedRequest.id,
        dialogAction === "approve" ? "onaylandi" : "reddedildi",
        approverComment
      );
      dispatchToast(
        <Toast>
          <ToastTitle>{dialogAction === "approve" ? "✅ Talep Onaylandı" : "❌ Talep Reddedildi"}</ToastTitle>
          <ToastBody>{selectedRequest.requesterName} adlı kullanıcıya bildirim gönderildi.</ToastBody>
        </Toast>,
        { intent: dialogAction === "approve" ? "success" : "warning" }
      );
      setSelectedRequest(null);
      setDialogAction(null);
      setApproverComment("");
      loadData();
    } catch (err) {
      dispatchToast(
        <Toast><ToastTitle>Hata</ToastTitle><ToastBody>{String(err)}</ToastBody></Toast>,
        { intent: "error" }
      );
    } finally {
      setIsProcessing(false);
    }
  };

  // ── Talep Kartı ──────────────────────────────────────────────────────────
  const RequestCard: React.FC<{ req: EnrichedLeaveRequest; showActions?: boolean }> = ({
    req, showActions = false,
  }) => {
    const isSelf = req.requesterId === currentUser.id;
    return (
      <Card className={req.status === "beklemede" ? styles.pendingCard : styles.requestCard}>
        <div className={styles.cardBody}>
          <div className={styles.metaRow}>
            <Avatar name={req.requesterName} size={32} color="colorful" />
            <div>
              <Text weight="semibold" size={300}>{req.requesterName}</Text>
              <br />
              <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
                {req.requesterEmail}
              </Text>
            </div>
            <Badge appearance="outline" color="brand">
              {LEAVE_TYPE_LABELS[req.leaveType]}
            </Badge>
            <StatusBadge status={req.status} />
            <Text size={200}>{req.totalDays} iş günü</Text>
            <VisibilityBadge meta={req._visibility} />
          </div>

          <Text size={300} weight="semibold">
            📅 {new Date(req.startDate).toLocaleDateString("tr-TR")} —{" "}
            {new Date(req.endDate).toLocaleDateString("tr-TR")}
          </Text>

          {req.description && (
            <Text size={200} style={{ color: "var(--colorNeutralForeground2)" }}>
              {req.description}
            </Text>
          )}

          {req.approverComment && (
            <div style={{
              padding: "8px 12px",
              background: "var(--colorNeutralBackground2)",
              borderRadius: "4px",
              borderLeft: "3px solid var(--colorNeutralStroke1)",
            }}>
              <Text size={200} weight="semibold">{req.approverName}: </Text>
              <Text size={200}>{req.approverComment}</Text>
            </div>
          )}

          <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
            {new Date(req.createdAt).toLocaleString("tr-TR")}
          </Text>

          {showActions && (
            <div className={styles.actions}>
              {isSelf ? (
                <Text size={100} style={{ color: "var(--colorNeutralForeground3)", fontStyle: "italic" }}>
                  Kendi talebinizi onaylayamazsınız
                </Text>
              ) : (
                <>
                  <Button
                    appearance="primary"
                    icon={<CheckmarkCircleRegular />}
                    onClick={() => { setSelectedRequest(req); setDialogAction("approve"); }}
                  >
                    Onayla
                  </Button>
                  <Button
                    appearance="subtle"
                    icon={<DismissCircleRegular />}
                    onClick={() => { setSelectedRequest(req); setDialogAction("reject"); }}
                  >
                    Reddet
                  </Button>
                </>
              )}
            </div>
          )}
        </div>
      </Card>
    );
  };

  if (isLoading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", padding: "48px" }}>
        <Spinner label="Yükleniyor..." />
      </div>
    );
  }

  return (
    <>
      <Toaster toasterId={toasterId} position="top-end" />

      {/* Onay/Red Dialog */}
      <Dialog
        open={!!selectedRequest && !!dialogAction}
        onOpenChange={(_, d) => {
          if (!d.open) { setSelectedRequest(null); setDialogAction(null); setApproverComment(""); }
        }}
      >
        <DialogSurface>
          <DialogTitle>
            {dialogAction === "approve" ? "✅ Talebi Onayla" : "❌ Talebi Reddet"}
          </DialogTitle>
          <DialogBody>
            {selectedRequest && (
              <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                <Text size={300}>
                  <strong>{selectedRequest.requesterName}</strong> adlı kullanıcının{" "}
                  <strong>{LEAVE_TYPE_LABELS[selectedRequest.leaveType]}</strong>{" "}
                  talebini {dialogAction === "approve" ? "onaylıyorsunuz" : "reddediyorsunuz"}.
                </Text>
                <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
                  {new Date(selectedRequest.startDate).toLocaleDateString("tr-TR")} —{" "}
                  {new Date(selectedRequest.endDate).toLocaleDateString("tr-TR")}{" "}
                  ({selectedRequest.totalDays} iş günü)
                </Text>
                <Field label={dialogAction === "reject" ? "Red Gerekçesi (zorunlu)" : "Yorum (opsiyonel)"}>
                  <Textarea
                    placeholder={
                      dialogAction === "approve"
                        ? "Onay notu ekleyebilirsiniz..."
                        : "Red gerekçesini belirtiniz..."
                    }
                    value={approverComment}
                    onChange={(_, d) => setApproverComment(d.value)}
                    rows={3}
                  />
                </Field>
              </div>
            )}
          </DialogBody>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary">İptal</Button>
            </DialogTrigger>
            <Button
              appearance={dialogAction === "approve" ? "primary" : "subtle"}
              icon={dialogAction === "approve" ? <CheckmarkCircleRegular /> : <DismissCircleRegular />}
              onClick={handleAction}
              disabled={isProcessing || (dialogAction === "reject" && !approverComment.trim())}
            >
              {isProcessing ? "İşleniyor..." : dialogAction === "approve" ? "Onayla" : "Reddet"}
            </Button>
          </DialogActions>
        </DialogSurface>
      </Dialog>

      <div className={styles.panel}>
        <TabList
          selectedValue={activeTab}
          onTabSelect={(_, d) => setActiveTab(d.value as AdminTab)}
        >
          {hasDirectReports && (
            <Tab value="pending" icon={<ClipboardTaskListLtrRegular />}>
              Bekleyen Talepler
              {pendingRequests.length > 0 && (
                <Badge appearance="filled" color="danger" size="small" style={{ marginLeft: "6px" }}>
                  {pendingRequests.length}
                </Badge>
              )}
            </Tab>
          )}
          <Tab value="all">Tüm Talepler</Tab>
          {isTreeAdmin && (
            <Tab value="hierarchy" icon={<PeopleRegular />}>Hiyerarşi Yönetimi</Tab>
          )}
        </TabList>

        {/* ── Bekleyen Talepler ─── */}
        {activeTab === "pending" && hasDirectReports && (
          <div>
            {/* Görünürlük notu */}
            <div className={styles.visibilityNote}>
              <LockClosedRegular fontSize={13} />
              <Text size={100}>
                Sadece doğrudan astlarınızın talepleri burada görünür. Onaylandıktan sonra çalışan kendi talebini görebilir.
              </Text>
            </div>
            <Toolbar>
              <ToolbarButton icon={<ArrowClockwiseRegular />} onClick={loadData}>Yenile</ToolbarButton>
            </Toolbar>
            {pendingRequests.length === 0 ? (
              <div className={styles.emptyState}>
                <CheckmarkCircleRegular fontSize={48} />
                <Text size={300}>Bekleyen talep bulunmuyor.</Text>
              </div>
            ) : (
              pendingRequests.map((req) => <RequestCard key={req.id} req={req} showActions />)
            )}
          </div>
        )}

        {/* ── Tüm Talepler ─── */}
        {activeTab === "all" && (
          <div>
            <div className={styles.filterBar}>
              <Text size={200}>Filtre:</Text>
              <Select
                value={filterStatus}
                onChange={(_, d) => setFilterStatus(d.value as LeaveStatus | "hepsi")}
                style={{ width: "160px" }}
              >
                <option value="hepsi">Tümü</option>
                <option value="beklemede">Beklemede</option>
                <option value="onaylandi">Onaylandı</option>
                <option value="reddedildi">Reddedildi</option>
              </Select>
              <ToolbarButton icon={<ArrowClockwiseRegular />} onClick={loadData}>Yenile</ToolbarButton>
            </div>
            {filteredRequests.length === 0 ? (
              <div className={styles.emptyState}><Text size={300}>Gösterilecek talep yok.</Text></div>
            ) : (
              <Table size="small" style={{ marginTop: "12px" }}>
                <TableHeader>
                  <TableRow>
                    <TableHeaderCell>Çalışan</TableHeaderCell>
                    <TableHeaderCell>Tür</TableHeaderCell>
                    <TableHeaderCell>Tarihler</TableHeaderCell>
                    <TableHeaderCell>Gün</TableHeaderCell>
                    <TableHeaderCell>Durum</TableHeaderCell>
                    <TableHeaderCell>Görünürlük</TableHeaderCell>
                    <TableHeaderCell>Onaylayan</TableHeaderCell>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredRequests.map((req) => (
                    <TableRow key={req.id}>
                      <TableCell>
                        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                          <Avatar name={req.requesterName} size={24} color="colorful" />
                          <Text size={200}>{req.requesterName}</Text>
                        </div>
                      </TableCell>
                      <TableCell>
                        <Badge appearance="outline" size="small">
                          {LEAVE_TYPE_LABELS[req.leaveType]}
                        </Badge>
                      </TableCell>
                      <TableCell>
                        <Text size={200}>
                          {new Date(req.startDate).toLocaleDateString("tr-TR")} —{" "}
                          {new Date(req.endDate).toLocaleDateString("tr-TR")}
                        </Text>
                      </TableCell>
                      <TableCell><Text size={200}>{req.totalDays}</Text></TableCell>
                      <TableCell><StatusBadge status={req.status} /></TableCell>
                      <TableCell>
                        {req._visibility?.badge !== "normal" ? (
                          <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", color: "var(--colorNeutralForeground3)", fontSize: "12px" }}>
                            <LockClosedRegular fontSize={11} /> Kısıtlı
                          </span>
                        ) : (
                          <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>Herkese Açık</Text>
                        )}
                      </TableCell>
                      <TableCell><Text size={200}>{req.approverName ?? "—"}</Text></TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            )}
          </div>
        )}

        {/* ── Hiyerarşi Yönetimi ─── */}
        {activeTab === "hierarchy" && isTreeAdmin && (
          <HierarchyEditor token={token} currentUserId={currentUser.id} />
        )}
      </div>
    </>
  );
};
