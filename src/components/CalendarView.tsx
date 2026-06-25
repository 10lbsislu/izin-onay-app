/**
 * CalendarView.tsx
 * "Takvim" sekmesi — HERKESE AÇIK.
 * BUGÜNDEN itibaren 30 günü gösterir (geçmiş gösterilmez).
 * Her güne o günkü onaylı izinler ve doğum günleri işaretlenir.
 */

import React, { useEffect, useState, useCallback, useMemo } from "react";
import {
  Text, Spinner, Button, Card, makeStyles, tokens,
  MessageBar, MessageBarBody, Badge,
} from "@fluentui/react-components";
import { ArrowClockwiseRegular } from "@fluentui/react-icons";
import type { CalendarLeave, Birthday } from "../services/requestService";
import { getCalendar } from "../services/requestService";
import { LEAVE_TYPE_LABELS } from "../types";

const WINDOW_DAYS = 30;

const MONTHS_SHORT = [
  "Oca", "Şub", "Mar", "Nis", "May", "Haz",
  "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara",
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
  rangeText: { color: tokens.colorNeutralForeground3 },
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
  dayHead: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  monthTag: { fontSize: "10px", fontWeight: tokens.fontWeightSemibold, color: tokens.colorBrandForeground1 },
  dayNum: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 },
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

function ymd(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function mmdd(d: Date): string {
  return `${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function addDays(base: Date, n: number): Date {
  const d = new Date(base);
  d.setDate(d.getDate() + n);
  return d;
}

interface CalendarViewProps { token: string; }

export const CalendarView: React.FC<CalendarViewProps> = ({ token }) => {
  const styles = useStyles();
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

  // Bugünden itibaren 30 gün + grid hizalaması (Pazartesi başlangıçlı)
  const { cells, todayStr, rangeLabel } = useMemo(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const days: Date[] = Array.from({ length: WINDOW_DAYS }, (_, i) => addDays(today, i));
    const startDow = (today.getDay() + 6) % 7; // Pzt=0
    const arr: (Date | null)[] = [];
    for (let i = 0; i < startDow; i++) arr.push(null);
    days.forEach((d) => arr.push(d));
    while (arr.length % 7 !== 0) arr.push(null);
    const last = days[days.length - 1];
    const range = `${today.getDate()} ${MONTHS_SHORT[today.getMonth()]} – ${last.getDate()} ${MONTHS_SHORT[last.getMonth()]} ${last.getFullYear()}`;
    return { cells: arr, todayStr: ymd(today), rangeLabel: range };
  }, []);

  const leavesOnDay = (dateStr: string): CalendarLeave[] =>
    leaves.filter((l) => l.startDate && l.endDate && dateStr >= l.startDate && dateStr <= l.endDate);

  const birthdaysOnDate = (d: Date): Birthday[] => {
    const key = mmdd(d);
    return birthdays.filter((b) => String(b.birthDate) === key);
  };

  // Penceredeki olay sayıları (özet)
  const summary = useMemo(() => {
    const dates = cells.filter((c): c is Date => c !== null);
    const bd = birthdays.filter((b) => dates.some((d) => mmdd(d) === String(b.birthDate))).length;
    const lv = leaves.filter((l) => dates.some((d) => { const s = ymd(d); return s >= l.startDate && s <= l.endDate; })).length;
    return { bd, lv };
  }, [cells, birthdays, leaves]);

  if (isLoading) {
    return <div style={{ display: "flex", justifyContent: "center", padding: "48px" }}>
      <Spinner label="Takvim yükleniyor..." /></div>;
  }

  return (
    <div className={styles.panel}>
      <div className={styles.header}>
        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.3 }}>
          <Text weight="semibold" size={500}>Takvim — Önümüzdeki 30 Gün</Text>
          <Text size={200} className={styles.rangeText}>{rangeLabel}</Text>
        </div>
        <Button appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={load} title="Yenile">Yenile</Button>
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
            const dateStr = ymd(d);
            const dayLeaves = leavesOnDay(dateStr);
            const dayBirthdays = birthdaysOnDate(d);
            const isToday = dateStr === todayStr;
            const showMonth = d.getDate() === 1 || dateStr === todayStr;
            return (
              <div key={dateStr} className={styles.dayCell}>
                <div className={styles.dayHead}>
                  <span className={styles.monthTag}>{showMonth ? MONTHS_SHORT[d.getMonth()] : ""}</span>
                  <span className={isToday ? styles.todayNum : styles.dayNum}>{d.getDate()}</span>
                </div>
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
        Bu pencerede: <Badge appearance="tint" size="small">{summary.lv}</Badge> izin ·{" "}
        <Badge appearance="tint" size="small">{summary.bd}</Badge> doğum günü
      </Text>
    </div>
  );
};
