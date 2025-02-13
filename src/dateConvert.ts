export function parseExcelDate(excelTimestamp: number): Date {
    const baseDate = new Date(1900, 0, 1);
    baseDate.setDate(baseDate.getDate() + excelTimestamp - 2);
    return baseDate
}