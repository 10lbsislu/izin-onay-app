/**
 * ApprovedLeaves.tsx
 * "Onaylanan İzinler" sekmesi — yetkili (isApprovalViewer) kullanıcılar görür.
 * Kuruluştaki TÜM onaylanan izinleri listeler ve Excel'e (.xlsx) aktarır.
 */

import React, { useEffect, useState, useCallback } from "react";
import {
  Text, Spinner, Button, Card, Badge, Avatar,
  Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell,
  Input, makeStyles, tokens,
  MessageBar, MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowClockwiseRegular, ArrowDownloadRegular,
  CheckmarkCircleRegular, SearchRegular,
} from "@fluentui/react-icons";
import * as XLSX from "xlsx";
import type { LeaveRequest } from "../types";
import { LEAVE_TYPE_LABELS, formatDuration } from "../types";
import { getApprovedLeaves } from "../services/requestService";

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", gap: "16px" },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: "12px",
    padding: "12px 16px",
    background: `linear-gradient(135deg, ${tokens.colorPaletteGreenBackground1} 0%, ${tokens.colorBrandBackground2} 100%)`,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  toolbar: { display: "flex", gap: "8px", alignItems: "center", flexWrap: "wrap" },
  approvedBy: { color: tokens.colorPaletteGreenForeground1, fontWeight: tokens.fontWeightSemibold },
  emptyState: {
    display: "flex", flexDirection: "column", alignItems: "center",
    gap: "12px", padding: "56px 24px", color: tokens.colorNeutralForeground3,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
  },
});

function fmtDate(d: string): string {
  if (!d) return "—";
  return new Date(d).toLocaleDateString("tr-TR");
}

function dateRange(req: LeaveRequest): string {
  if (req.leaveType === "saatlik" && req.startTime && req.endTime) {
    return `${fmtDate(req.startDate)} ${req.startTime}–${req.endTime}`;
  }
  if (req.startDate === req.endDate) return fmtDate(req.startDate);
  return `${fmtDate(req.startDate)} — ${fmtDate(req.endDate)}`;
}

interface ApprovedLeavesProps {
  token: string;
}

export const ApprovedLeaves: React.FC<ApprovedLeavesProps> = ({ token }) => {
  const styles = useStyles();
  const [requests, setRequests] = useState<LeaveRequest[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState("");

  const load = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const data = await getApprovedLeaves(token);
      setRequests(data);
    } catch (err) {
      setError(String(err));
    } finally {
      setIsLoading(false);
    }
  }, [token]);

  useEffect(() => { load(); }, [load]);

  const filtered = search.trim()
    ? requests.filter((r) => {
        const q = search.toLocaleLowerCase("tr");
        return (
          r.requesterName?.toLocaleLowerCase("tr").includes(q) ||
          (r.approverName || "").toLocaleLowerCase("tr").includes(q)
        );
      })
    : requests;

  // ── Excel'e aktar ──────────────────────────────────────────────────────────
  const exportToExcel = () => {
    const rows = filtered.map((r) => ({
      "Çalışan": r.requesterName,
      "E-posta": r.requesterEmail,
      "İzin Türü": LEAVE_TYPE_LABELS[r.leaveType] ?? r.leaveType,
      "Başlangıç": fmtDate(r.startDate),
      "Bitiş": r.leaveType === "saatlik" ? fmtDate(r.startDate) : fmtDate(r.endDate),
      "Saat": r.leaveType === "saatlik" && r.startTime && r.endTime ? `${r.startTime}–${r.endTime}` : "",
      "Süre": formatDuration(r),
      "Onaylayan": r.approverName ?? "",
      "Onay Tarihi": r.updatedAt ? new Date(r.updatedAt).toLocaleString("tr-TR") : "",
      "Açıklama": r.description ?? "",
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    // Kolon genişlikleri
    ws["!cols"] = [
      { wch: 22 }, { wch: 28 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
      { wch: 12 }, { wch: 16 }, { wch: 22 }, { wch: 20 }, { wch: 36 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Onaylanan İzinler");

    const today = new Date().toLocaleDateString("tr-TR").replace(/\./g, "-");
    XLSX.writeFile(wb, `onaylanan-izinler-${today}.xlsx`);
  };

  if (isLoading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", padding: "48px" }}>
        <Spinner label="Onaylanan izinler yükleniyor..." />
      </div>
    );
  }

  return (
    <div className={styles.panel}>
      <div className={styles.header}>
        <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
          <CheckmarkCircleRegular fontSize={28} />
          <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.2 }}>
            <Text weight="semibold" size={400}>Tüm Onaylanan İzinler</Text>
            <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
              Kuruluş genelinde onaylanmış {requests.length} izin
            </Text>
          </div>
        </div>
        <div className={styles.toolbar}>
          <Input
            placeholder="Çalışan / onaylayan ara..."
            value={search}
            onChange={(_, d) => setSearch(d.value)}
            contentBefore={<SearchRegular />}
            style={{ minWidth: "220px" }}
          />
          <Button icon={<ArrowClockwiseRegular />} appearance="subtle" onClick={load}>
            Yenile
          </Button>
          <Button
            icon={<ArrowDownloadRegular />}
            appearance="primary"
            onClick={exportToExcel}
            disabled={filtered.length === 0}
          >
            Excel'e Aktar
          </Button>
        </div>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!error && filtered.length === 0 ? (
        <div className={styles.emptyState}>
          <CheckmarkCircleRegular fontSize={56} />
          <Text size={400} weight="semibold">
            {requests.length === 0 ? "Henüz onaylanan izin yok" : "Aramayla eşleşen kayıt yok"}
          </Text>
        </div>
      ) : !error && (
        <Card>
          <Table size="small">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Çalışan</TableHeaderCell>
                <TableHeaderCell>Tür</TableHeaderCell>
                <TableHeaderCell>Tarihler</TableHeaderCell>
                <TableHeaderCell>Süre</TableHeaderCell>
                <TableHeaderCell>Onaylayan</TableHeaderCell>
                <TableHeaderCell>Onay Tarihi</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {filtered.map((req) => (
                <TableRow key={req.id}>
                  <TableCell>
                    <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                      <Avatar name={req.requesterName} size={24} color="colorful" />
                      <Text size={200}>{req.requesterName}</Text>
                    </div>
                  </TableCell>
                  <TableCell>
                    <Badge appearance="outline" size="small">
                      {LEAVE_TYPE_LABELS[req.leaveType] ?? req.leaveType}
                    </Badge>
                  </TableCell>
                  <TableCell><Text size={200}>{dateRange(req)}</Text></TableCell>
                  <TableCell><Text size={200}>{formatDuration(req)}</Text></TableCell>
                  <TableCell>
                    <Text size={200} className={styles.approvedBy}>
                      {req.approverName || "—"}
                    </Text>
                  </TableCell>
                  <TableCell>
                    <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
                      {req.updatedAt ? new Date(req.updatedAt).toLocaleString("tr-TR") : "—"}
                    </Text>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </Card>
      )}
    </div>
  );
};
