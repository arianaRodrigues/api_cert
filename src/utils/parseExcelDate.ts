export const parseExcelDate = (value: any): Date | null => {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === "number") {
    const excelEpoch = new Date(1900, 0, 1);
    const date = new Date(excelEpoch.getTime() + (value - 2) * 86400000);
    return isNaN(date.getTime()) ? null : date;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();

    const match = trimmed.match(/^(\d{2})\.(\d{2})\.(\d{2,4})$/);
    if (match) {
      const [_, day, month, year] = match;
      const fullYear = year.length === 2 ? `20${year}` : year;
       const date = new Date(`${fullYear}-${month}-${day}`);
      return isNaN(date.getTime()) ? null : date;
    }

    const match2 = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (match2) {
      const [_, day, month, year] = match2;
      const date = new Date(`${year}-${month}-${day}`);
      return isNaN(date.getTime()) ? null : date;
    }

    const parsed = new Date(trimmed);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
};
