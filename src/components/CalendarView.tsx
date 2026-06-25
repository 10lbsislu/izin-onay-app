/**
 * CalendarView.tsx
 * "Takvim" sekmesi — HERKESE AÇIK.
 * Seçili 1 ayı grid olarak gösterir; o aydaki onaylı izinler ve doğum günleri işaretlenir.
 */

import React, { useEffect, useState, useCallback, useMemo } from "react";
import {
  Text, Spinner, Button, Card, makeStyles, tokens,
  MessageBar, MessageBarBody, Badge,
} from "@fluentui/react-components";
import {
  ChevronLeftRegular, ChevronRightRegular, ArrowClockwiseRegular,
} from "@fluentui/react-icons";
import type { CalendarLeave, Birthday } from "../services/requestService";
import { getCalendar } from "../services/requestService";
import { LEAVE_TYPE_LABELS } from "../types";

const MONTHS = [
  "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık",
];
const WEEKDAYS = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"];

const LEAVE_COLOR: Record<string, string> = {
  yillik: tokens.colorPaletteBlueBorderActive,
  saatlik: tokens.colorPaletteTealBorderActive,
  ucretsiz: tokens.colorPaletteBeigeBorderActive,
  diger: tokens.colorPalettePurpleBorderActive,
};

const useStyles = makeStyles({
  panel: { display: "flex", flexDirection: "column", gap: "12px", maxWidth: "1000px", margin: "0 auto" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" },
  navGroup: { display: "flex", alignItems: "center", gap: "8px" },
  monthTitle: { minWidth: "160px", textAlign: "center" },
  legend: { display: "flex", gap: "12px", flexWrap: "wrap", alignItems: "center", padding: "4px 2px" },
  legendItem: { display: "flex", alignItems: "center", gap: "6px" },
  dot: { width: "10px", height: "10px", borderRadius: "50%", display: "inline-block" },
  grid: {
    display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: "1px",
    backgroundColor: tokens.colorNeutralStroke2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium, overflow: "hidden",
  },
  weekdayCell: {
    backgroundColor: tokens.colorNeutralBackground3, padding: "8px 6px",
    textAlign: "center", fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase200,
  },
  dayCell: {
    backgroundColor: tokens.colorNeutralBackground1, minHeight: "104px",
    padding: "4px 5px", display: "flex", flexDirection: "column", gap: "3px",
  },
  dayCellMuted: { backgroundColor: tokens.colorNeutralBackground2 },
  dayNum: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, alignSelf: "flex-end" },
  todayNum: {
    backgroundColor: tokens.colorBrandBackground, color: tokens.colorNeutralForegroundOnBrand,
    borderRadius: "50%", width: "20px", height: "20px", display: "flex",
    alignItems: "center", justifyContent: "center", fontWeight: tokens.fontWeightSemibold,
  },
  event: {
    fontSize: "11px", padding: "1px 5px", borderRadius: tokens.borderRadiusSmall,
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
    borderLeft: "3px solid", backgroundColor: tokens.colorNeutralBackground3,
  },
  birthday: {
    fontSize: "11px", padding: "1px 5px", borderRadius: tokens.borderRadiusSmall,
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
    backgroundColor: tokens.colorPalettePinkBackground2, color: tokens.colorPaletteMagentaForeground2,
  },
});

function ymd(y: number, m: number, d: number): string {
  return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

interface CalendarViewProps { token: string; }

export const CalendarView: React.FC<CalendarViewProps> = ({ token }) => {
  const styles = useStyles();
  const now = new Date();
  const [cursor, setCursor] = useState<{ y: number; m: number }>({ y: now.getFullYear(), m: now.getMonth() });
  const [leaves, setLeaves] = useState<CalendarLeave[]>([]);
  const [birthdays, setBirthdays] = useState<Birthday[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setIsLoading(true); setError(null);
    try {
      const data = await getCalendar(token);
      setLeaves(data.leaves);
      setBirthdays(data.birthdays);
    } catch (err) {
      setError(String(err));
    } finally {
      setIsLoading(false);
    }
  }, [token]);

  useEffect(() => { load(); }, [load]);

  const todayStr = ymd(now.getFullYear(), now.getMonth(), now.getDate());

  // Ayın günlerini ve baştaki boşlukları hesapla (Pazartesi başlangıçlı)
  const cells = useMemo(() => {
    const firstDow = (new Date(cursor.y, cursor.m, 1).getDay() + 6) % 7; // Pzt=0
    const daysInMonth = new Date(cursor.y, cursor.m + 1, 0).getDate();
    const arr: (number | null)[] = [];
    for (let i = 0; i < firstDow; i++) arr.push(null);
    for (let d = 1; d <= daysInMonth; d++) arr.push(d);
    while (arr.length % 7 !== 0) arr.push(null);
    return arr;
  }, [cursor]);

  const leavesOnDay = (dateStr: string): CalendarLeave[] =>
    leaves.filter((l) => l.startDate && l.endDate && dateStr >= l.startDate && dateStr <= l.endDate);

  const mmdd = (m: number, d: number) => `${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  const birthdaysOnDay = (d: number): Birthday[] =>
    birthdays.filter((b) => b.birthDate === mmdd(cursor.m, d));

  const prevMonth = () => setCursor((c) => (c.m === 0 ? { y: c.y - 1, m: 11 } : { y: c.y, m: c.m - 1 }));
  const nextMonth = () => setCursor((c) => (c.m === 11 ? { y: c.y + 1, m: 0 } : { y: c.y, m: c.m + 1 }));
  const goToday = () => setCursor({ y: now.getFullYear(), m: now.getMonth() });

  if (isLoading) {
    return <div style={{ display: "flex", justifyContent: "center", padding: "48px" }}>
      <Spinner label="Takvim yükleniyor..." /></div>;
  }

  return (
    <div className={styles.panel}>
      <div className={styles.header}>
        <Text weight="semibold" size={500}>Takvim</Text>
        <div className={styles.navGroup}>
          <Button appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={load} title="Yenile" />
          <Button appearance="subtle" size="small" onClick={goToday}>Bugün</Button>
          <Button appearance="subtle" icon={<ChevronLeftRegular />} onClick={prevMonth} title="Önceki ay" />
          <Text weight="semibold" size={400} className={styles.monthTitle}>{MONTHS[cursor.m]} {cursor.y}</Text>
          <Button appearance="subtle" icon={<ChevronRightRegular />} onClick={nextMonth} title="Sonraki ay" />
        </div>
      </div>

      <div className={styles.legend}>
        {(Object.keys(LEAVE_TYPE_LABELS) as (keyof typeof LEAVE_TYPE_LABELS)[]).map((k) => (
          <span key={k} className={styles.legendItem}>
            <span className={styles.dot} style={{ backgroundColor: LEAVE_COLOR[k] }} />
            <Text size={100}>{LEAVE_TYPE_LABELS[k]}</Text>
          </span>
        ))}
        <span className={styles.legendItem}>
          <span style={{ fontSize: 13 }}>🎂</span><Text size={100}>Doğum Günü</Text>
        </span>
      </div>

      {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}

      <Card style={{ padding: 0 }}>
        <div className={styles.grid}>
          {WEEKDAYS.map((w) => <div key={w} className={styles.weekdayCell}>{w}</div>)}
          {cells.map((d, i) => {
            if (d === null) return <div key={`e${i}`} className={`${styles.dayCell} ${styles.dayCellMuted}`} />;
            const dateStr = ymd(cursor.y, cursor.m, d);
            const dayLeaves = leavesOnDay(dateStr);
            const dayBirthdays = birthdaysOnDay(d);
            const isToday = dateStr === todayStr;
            return (
              <div key={dateStr} className={styles.dayCell}>
                <span className={isToday ? styles.todayNum : styles.dayNum}>{d}</span>
                {dayBirthdays.map((b) => (
                  <span key={b.id} className={styles.birthday} title={`🎂 ${b.name}`}>🎂 {b.name}</span>
                ))}
                {dayLeaves.map((l) => (
                  <span key={l.id} className={styles.event}
                    style={{ borderLeftColor: LEAVE_COLOR[l.leaveType] || tokens.colorNeutralStroke1 }}
                    title={`${l.requesterName} — ${LEAVE_TYPE_LABELS[l.leaveType]}${l.leaveType === "saatlik" && l.startTime ? ` (${l.startTime}–${l.endTime})` : ""}`}>
                    {l.requesterName}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
      </Card>

      <Text size={100} style={{ color: "var(--colorNeutralForeground3)" }}>
        <Badge appearance="tint" size="small">{leaves.length}</Badge> onaylı izin ·{" "}
        <Badge appearance="tint" size="small">{birthdays.length}</Badge> kayıtlı doğum günü
      </Text>
    </div>
  );
};
