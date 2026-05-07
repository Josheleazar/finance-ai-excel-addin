/*
 * Typed Excel helpers. All functions batch work inside a single Excel.run
 * and call context.sync() once to minimize round-trips.
 */

/* global Excel */

export type CellValue = string | number | boolean | null;
export type Row = CellValue[];

/**
 * Create a new worksheet (appending a numeric suffix if the name is taken),
 * write a table (header row + data rows) starting at A1, autofit columns,
 * and activate it.
 */
export async function writeTableToNewSheet(
  baseName: string,
  header: string[],
  rows: Row[]
): Promise<string> {
  let finalName = sanitizeSheetName(baseName);

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const existing = new Set(sheets.items.map((s) => s.name));
    if (existing.has(finalName)) {
      let i = 2;
      while (existing.has(`${finalName} (${i})`)) i++;
      finalName = `${finalName} (${i})`;
    }

    const sheet = sheets.add(finalName);

    // Header
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, header.length);
    headerRange.values = [header];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#f3f2f1";

    // Data
    if (rows.length > 0) {
      const dataRange = sheet.getRangeByIndexes(1, 0, rows.length, header.length);
      dataRange.values = rows;
    }

    const usedRange = sheet.getRangeByIndexes(0, 0, rows.length + 1, header.length);
    usedRange.format.autofitColumns();

    sheet.activate();
    await context.sync();
  });

  return finalName;
}

/**
 * Write a single key/value pair block (two columns: label, value) to a new sheet.
 */
export async function writeKeyValueToNewSheet(
  baseName: string,
  title: string,
  pairs: Array<[string, CellValue]>
): Promise<string> {
  const header = [title, ""];
  const rows: Row[] = pairs.map(([k, v]) => [k, v]);
  return writeTableToNewSheet(baseName, header, rows);
}

function sanitizeSheetName(name: string): string {
  // Excel: max 31 chars; cannot contain \ / ? * [ ] :
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_").trim();
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned || "Sheet";
}
