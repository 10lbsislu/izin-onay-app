/**
 * MyRequests.tsx
 * Çalışanın kendi taleplerini listeler.
 * Beklemedeki taleplerde görünürlük rozeti gösterilir.
 */

import React, { useEffect, useState } from "react";
import {
  Card,
  Text,
  Spinner,
  Badge,
  Button,
  makeStyles,
  tokens,
  Divider,
  MessageBar,
  MessageBarBody,
  InfoLabel,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular, DocumentRegular } from "@fluentui/react-icons";
import type { EnrichedLeaveRequest, GetRequestsResponse } from "../services/requestService";
import { getRequests } from "../services/requestService";
import { LEAVE_TYPE_LABELS } from "../types";
import type { OrgUser } from "../types";
import { StatusBadge } from "./StatusBadge";
import { VisibilityBadge } from "./VisibilityBadge";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    maxWidth: "720px",
    margin: "0 auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  requestCard: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  cardBody: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "4px 0",
  },
  metaRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: "10px",
    alignItems: "center",
  },
  commentBox: {
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    borderLeft: `3px solid ${tokens.colorNeutralStroke1}`,
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    padding: "48px 24px",
    color: tokens.colorNeutralForeground3,
  },
  pendingCard: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderLeft: `3px solid ${tokens.colorPaletteYellowBorderActive}`,
  },
});

interface MyRequestsProps {
  currentUser: OrgUser;
  token: string;
  refreshKey: number;
}

export const MyRequests: React.FC<MyRequestsProps> = ({
  currentUser,
  token,
  refreshKey,
}) => {
  const styles = useStyles();
  const [requests, setRequests] = useState<EnrichedLeaveRequest[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const load = async () => {
    setIsLoading(true);
    setError(null);
    try {
      // Sunucu görünürlük kurallarını uygular — sadece yetkili talepler gelir
      const res: GetRequestsResponse = await getRequests(token);
      // Sadece kendi taleplerini göster (onaylayıcı değilse zaten sadece kendisi gelir)
      const mine = res.requests.filter(
        (r) => r.requesterId === currentUser.id
      );
      setRequests(mine.sort((a, b) => b.createdAt.localeCompare(a.createdAt)));
    } catch (err) {
      setError(String(err));
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refreshKey]);

  if (isLoading) {
    return (
      <div style={{ display: "flex", justifyContent: "center", padding: "48px" }}>
        <Spinner label="Talepler yükleniyor..." />
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text weight="semibold" size={500}>İzin Taleplerim</Text>
        <Button appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={load}>
          Yenile
        </Button>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!error && requests.length === 0 && (
        <div className={styles.emptyState}>
          <DocumentRegular fontSize={48} />
          <Text size={300}>Henüz izin talebiniz bulunmuyor.</Text>
        </div>
      )}

      {requests.map((req) => (
        <Card
          key={req.id}
          className={req.status === "beklemede" ? styles.pendingCard : styles.requestCard}
        >
          <div className={styles.cardBody}>
            {/* Başlık satırı */}
            <div className={styles.metaRow}>
              <Badge appearance="outline" color="brand">
                {LEAVE_TYPE_LABELS[req.leaveType]}
              </Badge>
              <StatusBadge status={req.status} />
              <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
                {req.totalDays} iş günü
              </Text>
              {/* 🔒 Görünürlük rozeti — sadece beklemedeyken */}
              <VisibilityBadge meta={req._visibility} />
            </div>

            {/* Tarihler */}
            <Text size={300} weight="semibold">
              {new Date(req.startDate).toLocaleDateString("tr-TR")}
              {req.leaveType === "saatlik" && req.startTime && req.endTime
                ? ` (${req.startTime}–${req.endTime})`
                : req.startDate !== req.endDate
                ? ` — ${new Date(req.endDate).toLocaleDateString("tr-TR")}`
                : ""}
            </Text>

            {/* Açıklama */}
            {req.description && (
              <Text size={200} style={{ color: "var(--colorNeutralForeground2)" }}>
                {req.description}
              </Text>
            )}

            {/* Onaylayıcı yorumu */}
            {req.approverComment && (
              <>
                <Divider />
                <div className={styles.commentBox}>
                  <Text size={200} weight="semibold">
                    {req.approverName} — yorum:
                  </Text>
                  <br />
                  <Text size={200}>{req.approverComment}</Text>
                </div>
              </>
            )}

            {/* Alt meta */}
            <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
              Oluşturulma: {new Date(req.createdAt).toLocaleString("tr-TR")}
              {req.updatedAt &&
                ` · Güncelleme: ${new Date(req.updatedAt).toLocaleString("tr-TR")}`}
            </Text>
          </div>
        </Card>
      ))}
    </div>
  );
};
