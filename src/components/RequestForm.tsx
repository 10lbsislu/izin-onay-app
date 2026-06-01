/**
 * RequestForm.tsx
 * Çalışanın izin talebi oluşturduğu form bileşeni.
 */

import React, { useState, useEffect } from "react";
import {
  Button,
  Field,
  Input,
  Select,
  Textarea,
  Text,
  Toast,
  ToastTitle,
  ToastBody,
  useToastController,
  useId,
  Toaster,
  makeStyles,
  tokens,
  Card,
  CardHeader,
  Divider,
  Badge,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { CalendarLtrRegular, SendRegular } from "@fluentui/react-icons";
import type { OrgUser, LeaveType, HierarchyNode } from "../types";
import { createRequest, getHierarchy } from "../services/requestService";
import { LEAVE_TYPE_LABELS } from "../types";

const useStyles = makeStyles({
  card: {
    maxWidth: "640px",
    margin: "0 auto",
  },
  form: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "8px 0",
  },
  row: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "16px",
    "@media (max-width: 480px)": {
      gridTemplateColumns: "1fr",
    },
  },
  dateInfo: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
});

interface RequestFormProps {
  currentUser: OrgUser;
  token: string;
  isRoot: boolean;
  onSuccess: () => void;
}

export const RequestForm: React.FC<RequestFormProps> = ({
  currentUser,
  token,
  isRoot,
  onSuccess,
}) => {
  const styles = useStyles();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const [leaveType, setLeaveType] = useState<LeaveType>("yillik");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [description, setDescription] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [approverChoices, setApproverChoices] = useState<HierarchyNode[]>([]);
  const [pickedApproverId, setPickedApproverId] = useState("");
  const [startTime, setStartTime] = useState("09:00");
  const [endTime, setEndTime] = useState("18:00");

  const isHourly = leaveType === "saatlik";

  // 24-saatlik, 30 dk aralıklı saat seçenekleri (Türkiye standardı, AM/PM yok)
  const TIME_OPTIONS = React.useMemo(() => {
    const opts: string[] = [];
    for (let h = 0; h < 24; h++) {
      for (let m = 0; m < 60; m += 30) {
        opts.push(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
      }
    }
    return opts;
  }, []);

  // Root kullanıcı: hiyerarşiden olası onaylayıcıları yükle
  useEffect(() => {
    if (!isRoot) return;
    getHierarchy(token)
      .then((nodes) => {
        const choices = nodes.filter((n) => n.id !== currentUser.id);
        setApproverChoices(choices);
      })
      .catch(() => setApproverChoices([]));
  }, [isRoot, token, currentUser.id]);

  const calcDays = (): number => {
    if (isHourly) {
      if (!startTime || !endTime) return 0;
      const [sh, sm] = startTime.split(":").map(Number);
      const [eh, em] = endTime.split(":").map(Number);
      const hours = (eh + em / 60) - (sh + sm / 60);
      if (hours <= 0) return 0;
      return Math.round((hours / 8) * 100) / 100; // 2 ondalık
    }
    if (!startDate || !endDate) return 0;
    const start = new Date(startDate);
    const end = new Date(endDate);
    if (end < start) return 0;
    let count = 0;
    const cur = new Date(start);
    while (cur <= end) {
      const day = cur.getDay();
      if (day !== 0 && day !== 6) count++; // Hafta içi
      cur.setDate(cur.getDate() + 1);
    }
    return count;
  };

  const totalDays = calcDays();
  const hourlyHours = isHourly && startTime && endTime
    ? Math.max(0, (Number(endTime.split(":")[0]) + Number(endTime.split(":")[1]) / 60)
                  - (Number(startTime.split(":")[0]) + Number(startTime.split(":")[1]) / 60))
    : 0;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!startDate) {
      dispatchToast(
        <Toast><ToastTitle>Eksik Alan</ToastTitle><ToastBody>Tarih seçiniz.</ToastBody></Toast>,
        { intent: "error" }
      );
      return;
    }
    if (isHourly) {
      if (!startTime || !endTime) {
        dispatchToast(
          <Toast><ToastTitle>Eksik Alan</ToastTitle><ToastBody>Saat başlangıç ve bitiş giriniz.</ToastBody></Toast>,
          { intent: "error" }
        );
        return;
      }
      if (totalDays === 0) {
        dispatchToast(
          <Toast><ToastTitle>Geçersiz Saat</ToastTitle><ToastBody>Bitiş saati başlangıçtan büyük olmalı.</ToastBody></Toast>,
          { intent: "error" }
        );
        return;
      }
    } else {
      if (!endDate) {
        dispatchToast(
          <Toast><ToastTitle>Eksik Alan</ToastTitle><ToastBody>Bitiş tarihini giriniz.</ToastBody></Toast>,
          { intent: "error" }
        );
        return;
      }
      if (new Date(endDate) < new Date(startDate)) {
        dispatchToast(
          <Toast><ToastTitle>Tarih Hatası</ToastTitle><ToastBody>Bitiş tarihi başlangıçtan önce olamaz.</ToastBody></Toast>,
          { intent: "error" }
        );
        return;
      }
      if (totalDays === 0) {
        dispatchToast(
          <Toast><ToastTitle>Geçersiz Tarih</ToastTitle><ToastBody>Seçilen tarihlerde iş günü bulunmuyor.</ToastBody></Toast>,
          { intent: "error" }
        );
        return;
      }
    }
    if (isRoot && !pickedApproverId) {
      dispatchToast(
        <Toast>
          <ToastTitle>Onaylayıcı Seçilmedi</ToastTitle>
          <ToastBody>Talebinizi onaylayacak kişiyi seçin.</ToastBody>
        </Toast>,
        { intent: "error" }
      );
      return;
    }

    setIsSubmitting(true);
    try {
      await createRequest(token, {
        requesterId: currentUser.id,
        requesterName: currentUser.displayName,
        requesterEmail: currentUser.mail,
        leaveType,
        startDate,
        endDate: isHourly ? startDate : endDate,
        totalDays,
        description,
        ...(isHourly ? { startTime, endTime } : {}),
        ...(isRoot ? { approverId: pickedApproverId } : {}),
      });

      dispatchToast(
        <Toast>
          <ToastTitle>Talep Oluşturuldu</ToastTitle>
          <ToastBody>İzin talebiniz onaylayıcıya iletildi.</ToastBody>
        </Toast>,
        { intent: "success" }
      );

      // Formu sıfırla
      setLeaveType("yillik");
      setStartDate("");
      setEndDate("");
      setDescription("");
      setPickedApproverId("");
      onSuccess();
    } catch (err) {
      dispatchToast(
        <Toast>
          <ToastTitle>Hata</ToastTitle>
          <ToastBody>{String(err)}</ToastBody>
        </Toast>,
        { intent: "error" }
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <>
      <Toaster toasterId={toasterId} position="top-end" />
      <Card className={styles.card}>
        <CardHeader
          header={
            <Text weight="semibold" size={500}>
              Yeni İzin Talebi
            </Text>
          }
          description={
            <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
              {currentUser.displayName} olarak talep açıyorsunuz
            </Text>
          }
        />
        <Divider />

        <form onSubmit={handleSubmit} className={styles.form}>
          {/* İzin Türü */}
          <Field label="İzin Türü" required>
            <Select
              value={leaveType}
              onChange={(_, d) => setLeaveType(d.value as LeaveType)}
            >
              {(Object.entries(LEAVE_TYPE_LABELS) as [LeaveType, string][]).map(
                ([val, lbl]) => (
                  <option key={val} value={val}>
                    {lbl}
                  </option>
                )
              )}
            </Select>
          </Field>

          {/* Tarih Aralığı */}
          {isHourly ? (
            <>
              <Field label="Tarih" required>
                <Input
                  type="date"
                  value={startDate}
                  onChange={(_, d) => setStartDate(d.value)}
                  min={new Date().toISOString().split("T")[0]}
                  contentBefore={<CalendarLtrRegular />}
                />
              </Field>
              <div className={styles.row}>
                <Field label="Saat Başlangıç" required>
                  <Select
                    value={startTime}
                    onChange={(_, d) => setStartTime(d.value)}
                  >
                    {TIME_OPTIONS.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </Select>
                </Field>
                <Field label="Saat Bitiş" required>
                  <Select
                    value={endTime}
                    onChange={(_, d) => setEndTime(d.value)}
                  >
                    {TIME_OPTIONS.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </Select>
                </Field>
              </div>
            </>
          ) : (
            <div className={styles.row}>
              <Field label="Başlangıç Tarihi" required>
                <Input
                  type="date"
                  value={startDate}
                  onChange={(_, d) => setStartDate(d.value)}
                  min={new Date().toISOString().split("T")[0]}
                  contentBefore={<CalendarLtrRegular />}
                />
              </Field>
              <Field label="Bitiş Tarihi" required>
                <Input
                  type="date"
                  value={endDate}
                  onChange={(_, d) => setEndDate(d.value)}
                  min={startDate || new Date().toISOString().split("T")[0]}
                  contentBefore={<CalendarLtrRegular />}
                />
              </Field>
            </div>
          )}

          {/* Hesaplanan gün/saat */}
          {totalDays > 0 && (
            <div className={styles.dateInfo}>
              <CalendarLtrRegular />
              <Text size={200}>
                {isHourly ? (
                  <>Toplam <strong>{hourlyHours.toFixed(2)} saat</strong> ({totalDays} iş günü karşılığı)</>
                ) : (
                  <>Toplam <strong>{totalDays} iş günü</strong> izin talebi</>
                )}
              </Text>
              <Badge appearance="outline" color="brand">
                {LEAVE_TYPE_LABELS[leaveType]}
              </Badge>
            </div>
          )}

          {/* Açıklama */}
          <Field label="Açıklama">
            <Textarea
              placeholder="İzin sebebi veya ek bilgi girebilirsiniz..."
              value={description}
              onChange={(_, d) => setDescription(d.value)}
              resize="vertical"
              rows={3}
            />
          </Field>

          {/* Root kullanıcı için onaylayıcı seçimi */}
          {isRoot && (
            approverChoices.length === 0 ? (
              <MessageBar intent="warning">
                <MessageBarBody>
                  Hiyerarşide başka kullanıcı yok. Önce hiyerarşiye birini ekleyin, sonra talep açabilirsiniz.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <Field label="Onaylayıcı" required hint="Hiyerarşinin tepesindesiniz; talebinizi onaylayacak kişiyi seçin.">
                <Select
                  value={pickedApproverId}
                  onChange={(_, d) => setPickedApproverId(d.value)}
                >
                  <option value="">— Seçiniz —</option>
                  {approverChoices.map((n) => (
                    <option key={n.id} value={n.id}>
                      {n.displayName}
                      {n.isTreeAdmin ? " (admin)" : ""}
                    </option>
                  ))}
                </Select>
              </Field>
            )
          )}

          <Button
            appearance="primary"
            type="submit"
            icon={<SendRegular />}
            disabled={isSubmitting}
          >
            {isSubmitting ? "Gönderiliyor..." : "Talebi Gönder"}
          </Button>
        </form>
      </Card>
    </>
  );
};
