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

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!startDate || !endDate) {
      dispatchToast(
        <Toast>
          <ToastTitle>Eksik Alan</ToastTitle>
          <ToastBody>Başlangıç ve bitiş tarihlerini doldurunuz.</ToastBody>
        </Toast>,
        { intent: "error" }
      );
      return;
    }
    if (new Date(endDate) < new Date(startDate)) {
      dispatchToast(
        <Toast>
          <ToastTitle>Tarih Hatası</ToastTitle>
          <ToastBody>Bitiş tarihi başlangıçtan önce olamaz.</ToastBody>
        </Toast>,
        { intent: "error" }
      );
      return;
    }
    if (totalDays === 0) {
      dispatchToast(
        <Toast>
          <ToastTitle>Geçersiz Tarih</ToastTitle>
          <ToastBody>Seçilen tarihlerde iş günü bulunmuyor.</ToastBody>
        </Toast>,
        { intent: "error" }
      );
      return;
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
        endDate,
        totalDays,
        description,
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

          {/* Hesaplanan gün */}
          {totalDays > 0 && (
            <div className={styles.dateInfo}>
              <CalendarLtrRegular />
              <Text size={200}>
                Toplam{" "}
                <strong>{totalDays} iş günü</strong> izin talebi
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
