const INVALID_DATE = 'Invalid date format. Expected YYYY-MM-DD or Excel serial number.';
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH_MS = Date.UTC(1900, 0, 1);
// Excel's own range: serial 1 is 1900-01-01, serial 2958465 is 9999-12-31
const MIN_SERIAL = 1;
const MAX_SERIAL = 2958465;

function serialToTimestamp(serial: number): number | null {
  if (!Number.isFinite(serial) || serial < MIN_SERIAL || serial > MAX_SERIAL) return null;
  return Math.floor((EXCEL_EPOCH_MS + (serial - 2) * MS_PER_DAY) / 1000);
}

// Local date parsing (moved from utils.ts since only used here)
function parseDate(dateInput: string | number): {timestamp: number; error?: string} {
  if (typeof dateInput === 'number') {
    const timestamp = serialToTimestamp(dateInput);
    return timestamp === null ? {timestamp: 0, error: INVALID_DATE} : {timestamp};
  }
  const dateStr = dateInput.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    // Reject dates that don't exist (2024-13-01, 2024-02-30) instead of reading them as serials
    const dateObj = new Date(dateStr + 'T00:00:00.000Z');
    if (isNaN(dateObj.getTime()) || dateObj.toISOString().slice(0, 10) !== dateStr) {
      return {timestamp: 0, error: INVALID_DATE};
    }
    return {timestamp: Math.floor(dateObj.getTime() / 1000)};
  }
  // A serial number typed as text, e.g. "45292"
  if (/^\d+(\.\d+)?$/.test(dateStr)) {
    const serial = parseFloat(dateStr);
    const timestamp = serial > 1000 ? serialToTimestamp(serial) : null;
    if (timestamp !== null) return {timestamp};
  }
  return {timestamp: 0, error: INVALID_DATE};
}

export type NormalizeDateSuccess = {ok: true; date: Date};
export type NormalizeDateError = {ok: false; error: string};
export type NormalizeDateResult = NormalizeDateSuccess | NormalizeDateError;

/**
 * Normalize Excel date inputs (serial number or YYYY-MM-DD string) into a Date.
 * Discriminated union ensures type safety.
 */
export function normalizeDateInput(input: string | number): NormalizeDateResult {
  const parsed = parseDate(input);
  if (parsed.error) return {ok: false, error: parsed.error};
  return {ok: true, date: new Date(parsed.timestamp * 1000)};
}
