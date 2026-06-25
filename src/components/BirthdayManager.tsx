/**
 * BirthdayManager.tsx
 * Yönetici panelinde doğum günü yönetimi (tree admin).
 * Manuel ekleme: isim + ay/gün. Listeleme getCalendar üzerinden.
 */

import React, { useEffect, useState, useCallback } from "react";
import {
  Card, Text, Spinner, Button, Input, Select, Field, Badge,
  makeStyles, tokens, Divider,
  Toast, ToastTitle, ToastBody, useToastController, useId, Toaster,
} from "@fluentui/react-components";
import { PersonAddRegular, DeleteRegular, ArrowClockwiseRegular } from "@fluentui/react-icons";
import type { Birthday } from "../services/requestService";
import { getCalendar, addBirthday, removeBirthday } from "../services/requestService";

const MONTHS = [
  "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
];

const useStyles = makeStyles({
  addBar: { display: "flex", gap: "10px", alignItems: "flex-end", flexWrap: "wrap" },
  list: { display: "flex", flexDirection: "column", gap: "6px", marginTop: "12px" },
  row: {
    display: "flex", alignItems: "center", gap: "10px", padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium, border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  empty: { padding: "24px", textAlign: "center", color: tokens.colorNeutralForeground3 },
});

function fmtBirthday(mmdd: string): string {
  const [mm, dd] = mmdd.split("-").map(Number);
  if (!mm || !dd) return mmdd;
  return `${dd} ${MONTHS[mm - 1] ?? ""}`;
}

interface Props { token: string; }

export const BirthdayManager: React.FC<Props> = ({ token }) => {
  const styles = useStyles();
  const toasterId = useId("bday-toaster");
  const { dispatchToast } = useToastController(toasterId);

  const [birthdays, setBirthdays] = useState<Birthday[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [name, setName] = useState("");
  const [month, setMonth] = useState("01");
  const [day, setDay] = useState("01");
  const [saving, setSaving] = useState(false);

  const toast = (t: string, b: string, intent: "success" | "warning" | "error" = "success") =>
    dispatchToast(<Toast><ToastTitle>{t}</ToastTitle><ToastBody>{b}</ToastBody></Toast>, { intent });

  const load = useCallback(async () => {
    setIsLoading(true);
    try {
      const data = await getCalendar(token);
      const sorted = [...data.birthdays].sort((a, b) => a.birthDate.localeCompare(b.birthDate));
      setBirthdays(sorted);
    } catch (err) {
      toast("Yükleme Hatası", String(err), "error");
    } finally {
      setIsLoading(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token]);

  useEffect(() => { load(); }, [load]);

  const handleAdd = async () => {
    if (!name.trim()) { toast("Eksik Alan", "İsim giriniz.", "error"); return; }
    setSaving(true);
    try {
      await addBirthday(token, name.trim(), `${month}-${day}`);
      toast("Eklendi", `${name.trim()} doğum günü eklendi.`);
      setName("");
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    } finally {
      setSaving(false);
    }
  };

  const handleRemove = async (b: Birthday) => {
    if (!confirm(`${b.name} doğum günü kaydı silinsin mi?`)) return;
    try {
      await removeBirthday(token, b.id);
      toast("Silindi", `${b.name} kaldırıldı.`, "warning");
      load();
    } catch (err) {
      toast("Hata", String(err), "error");
    }
  };

  const dayOptions = Array.from({ length: 31 }, (_, i) => String(i + 1).padStart(2, "0"));

  return (
    <>
      <Toaster toasterId={toasterId} position="top-end" />
      <Card>
        <Text weight="semibold" size={400}>Doğum Günü Ekle</Text>
        <Divider style={{ margin: "8px 0 12px" }} />
        <div className={styles.addBar}>
          <Field label="İsim" style={{ flex: 1, minWidth: "200px" }}>
            <Input value={name} onChange={(_, d) => setName(d.value)} placeholder="Ad Soyad" />
          </Field>
          <Field label="Ay">
            <Select value={month} onChange={(_, d) => setMonth(d.value)} style={{ width: "130px" }}>
              {MONTHS.map((m, i) => (
                <option key={i} value={String(i + 1).padStart(2, "0")}>{m}</option>
              ))}
            </Select>
          </Field>
          <Field label="Gün">
            <Select value={day} onChange={(_, d) => setDay(d.value)} style={{ width: "80px" }}>
              {dayOptions.map((d) => <option key={d} value={d}>{Number(d)}</option>)}
            </Select>
          </Field>
          <Button appearance="primary" icon={<PersonAddRegular />} onClick={handleAdd} disabled={saving}>
            {saving ? "Ekleniyor..." : "Ekle"}
          </Button>
          <Button appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={load} title="Yenile" />
        </div>

        {isLoading ? (
          <div style={{ display: "flex", justifyContent: "center", padding: "24px" }}>
            <Spinner label="Yükleniyor..." />
          </div>
        ) : birthdays.length === 0 ? (
          <div className={styles.empty}><Text size={300}>Henüz doğum günü kaydı yok.</Text></div>
        ) : (
          <div className={styles.list}>
            {birthdays.map((b) => (
              <div key={b.id} className={styles.row}>
                <span style={{ fontSize: 18 }}>🎂</span>
                <Text weight="semibold" size={300} style={{ flex: 1 }}>{b.name}</Text>
                <Badge appearance="outline" color="brand">{fmtBirthday(b.birthDate)}</Badge>
                <Button size="small" appearance="subtle" icon={<DeleteRegular />}
                  onClick={() => handleRemove(b)} title="Sil" />
              </div>
            ))}
          </div>
        )}
      </Card>
    </>
  );
};
