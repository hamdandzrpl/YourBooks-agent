import ExcelJS from "exceljs";

/** Render a workspace dataset (columns + rows) into a simple formatted .xlsx. */
export async function buildWorkspaceXlsx(dataset) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WhatsApp Finance Agent";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("Data");
  worksheet.addRow(dataset.columns ?? []);
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9F1ED" } };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  for (const row of dataset.rows ?? []) worksheet.addRow(row.map((cell) => normalizeCell(cell)));
  worksheet.columns.forEach((column, index) => {
    const header = String(dataset.columns?.[index] ?? "");
    column.width = Math.min(52, Math.max(12, header.length + 6));
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function normalizeCell(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return value;
  const text = String(value);
  if (/^-?\d+(?:[.,]\d+)?$/.test(text)) {
    const numeric = Number(text.replace(",", "."));
    if (Number.isFinite(numeric) && text.length <= 15) return numeric;
  }
  return text;
}
