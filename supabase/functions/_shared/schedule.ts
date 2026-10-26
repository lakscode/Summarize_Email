// Mailbox processing schedules: when is the next run?
//   daily     - every day at `time`
//   alternate - every other day at `time`, counting from `anchor` (the local date the schedule was saved)
//   weekdays  - on the selected `days` (0 = Sunday ... 6 = Saturday) at `time`
// Times are wall-clock in the user's IANA time zone, so DST changes keep the same local time.

export const FREQUENCIES = ["daily", "alternate", "weekdays"] as const;
export type Frequency = typeof FREQUENCIES[number];

export interface Schedule {
  frequency: Frequency;
  days: number[];
  time: string; // "HH:MM"
  timezone: string;
  anchor: string; // "YYYY-MM-DD" in `timezone`
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
