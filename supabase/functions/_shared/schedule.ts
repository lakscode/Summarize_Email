// Mailbox processing schedules: when is the next run?
//   daily     - every day at `time`
//   alternate - every other day at `time`, counting from `anchor` (the local date the schedule was saved)
//   weekdays  - on the selected `days` (0 = Sunday ... 6 = Saturday) at `time`
//   interval  - every `interval` minutes / hours / days, starting at `anchor` + `time`. Minutes and hours are
//               fixed spans of real time; days keep the same local time across DST changes.
// Times are wall-clock in the user's IANA time zone, so DST changes keep the same local time.

export const FREQUENCIES = ["daily", "alternate", "weekdays", "interval"] as const;
export type Frequency = typeof FREQUENCIES[number];
export const INTERVAL_UNITS = ["minutes", "hours", "days"] as const;
export type IntervalUnit = typeof INTERVAL_UNITS[number];
// Shortest interval allowed; matches the pg_cron tick that calls the scheduler.
export const MIN_INTERVAL_MINUTES = 5;

export interface Schedule {
  frequency: Frequency;
  days: number[];
  time: string; // "HH:MM"
  timezone: string;
  anchor: string; // "YYYY-MM-DD" in `timezone`
  interval?: number;
  intervalUnit?: IntervalUnit;
}

interface LocalParts { year: number; month: number; day: number; hour: number; minute: number }

export function isValidTimeZone(tz: string): boolean {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

function localParts(date: Date, tz: string): LocalParts {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit",
  }).formatToParts(date);
  const get = (type: string) => Number(parts.find((p) => p.type === type)?.value);
  return { year: get("year"), month: get("month"), day: get("day"), hour: get("hour"), minute: get("minute") };
}

// Milliseconds the zone is ahead of UTC at `date`.
function offsetMs(date: Date, tz: string): number {
  const p = localParts(date, tz);
  const asUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute);
  return asUtc - Math.floor(date.getTime() / 60_000) * 60_000;
}

// The UTC instant of a local wall-clock time in `tz`.
function zonedToUtc(year: number, month: number, day: number, hour: number, minute: number, tz: string): Date {
  const guess = Date.UTC(year, month - 1, day, hour, minute);
  let result = guess - offsetMs(new Date(guess), tz);
  const corrected = guess - offsetMs(new Date(result), tz);
  if (corrected !== result) result = corrected;
  return new Date(result);
}

export function localDate(date: Date, tz: string): string {
  const p = localParts(date, tz);
  return `${p.year}-${String(p.month).padStart(2, "0")}-${String(p.day).padStart(2, "0")}`;
}

const dayNumber = (y: number, m: number, d: number) => Math.round(Date.UTC(y, m - 1, d) / 86_400_000);

// First scheduled run strictly after `after`.
export function nextRun(s: Schedule, after: Date): Date {
  const [hh, mm] = s.time.split(":").map(Number);
  const start = localParts(after, s.timezone);
  const [ay, am, ad] = s.anchor.split("-").map(Number);
  const anchorDay = dayNumber(ay, am, ad);
  if (s.frequency === "interval") return nextInterval(s, after, ay, am, ad, hh, mm);

  for (let i = 0; i <= 15; i++) {
    const cal = new Date(Date.UTC(start.year, start.month - 1, start.day + i));
    const y = cal.getUTCFullYear(), m = cal.getUTCMonth() + 1, d = cal.getUTCDate();
    const allowed =
      s.frequency === "daily" ||
      (s.frequency === "alternate" && (((dayNumber(y, m, d) - anchorDay) % 2) + 2) % 2 === 0) ||
      (s.frequency === "weekdays" && s.days.includes(cal.getUTCDay()));
    if (!allowed) continue;
    const at = zonedToUtc(y, m, d, hh, mm, s.timezone);
    if (at.getTime() > after.getTime()) return at;
  }
  throw new Error("Schedule has no upcoming run");
}

function nextInterval(s: Schedule, after: Date, ay: number, am: number, ad: number, hh: number, mm: number): Date {
  const n = Math.max(1, s.interval ?? 1);
  const first = zonedToUtc(ay, am, ad, hh, mm, s.timezone);
  if (first.getTime() > after.getTime()) return first;

  if (s.intervalUnit === "days") {
    const start = localParts(after, s.timezone);
    const elapsedDays = dayNumber(start.year, start.month, start.day) - dayNumber(ay, am, ad);
    for (let k = Math.max(0, Math.floor(elapsedDays / n)); ; k++) {
      const at = zonedToUtc(ay, am, ad + k * n, hh, mm, s.timezone);
      if (at.getTime() > after.getTime()) return at;
    }
  }

  const step = n * (s.intervalUnit === "minutes" ? 60_000 : 3_600_000);
  const k = Math.floor((after.getTime() - first.getTime()) / step) + 1;
  return new Date(first.getTime() + k * step);
}
