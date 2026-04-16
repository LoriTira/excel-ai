/* global Excel */

const FORMULA_PREFIX = "EXCELAI.AI(";
const NUMBER_RE = /^-?\d+(\.\d+)?$/;

function isExcelAIFormula(formula: unknown): boolean {
  if (typeof formula !== "string" || formula.length === 0) return false;
  const stripped = formula[0] === "=" ? formula.slice(1) : formula;
  return stripped.trimStart().toUpperCase().startsWith(FORMULA_PREFIX);
}

function coerce(value: unknown): string | number | boolean {
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value !== "string") return String(value ?? "");
  const trimmed = value.trim();
  if (NUMBER_RE.test(trimmed)) return Number(trimmed);
  if (/^true$/i.test(trimmed)) return true;
  if (/^false$/i.test(trimmed)) return false;
  return value;
}

function isErrorCell(value: unknown): boolean {
  if (typeof value !== "string") return false;
  if (value.startsWith("#ERROR")) return true;
  return (
    value === "#NAME?" ||
    value === "#VALUE!" ||
    value === "#REF!" ||
    value === "#DIV/0!" ||
    value === "#N/A" ||
    value === "#NULL!" ||
    value === "#NUM!"
  );
}

export interface FreezeResult {
  frozen: number;
  skipped: number;
}

export async function freezeWorkbook(): Promise<FreezeResult> {
  let frozen = 0;
  let skipped = 0;

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    for (const sheet of sheets.items) {
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["formulas", "values", "rowCount", "columnCount", "isNullObject"]);
      await context.sync();

      if (used.isNullObject) continue;

      const formulas = used.formulas;
      const values = used.values;
      const next: (string | number | boolean)[][] = [];
      let anyFrozen = false;

      for (let r = 0; r < formulas.length; r++) {
        const row: (string | number | boolean)[] = [];
        for (let c = 0; c < formulas[r].length; c++) {
          const f = formulas[r][c];
          const v = values[r][c];
          if (isExcelAIFormula(f)) {
            if (isErrorCell(v)) {
              skipped++;
              row.push(f as string);
            } else {
              frozen++;
              anyFrozen = true;
              row.push(coerce(v));
            }
          } else {
            // Preserve existing cell content (formula or literal) by writing
            // the `formulas` entry back — Office.js treats "="-prefixed strings
            // as formulas and other primitives as literal values.
            row.push((f ?? "") as string | number | boolean);
          }
        }
        next.push(row);
      }

      if (anyFrozen) {
        used.formulas = next;
        await context.sync();
      }
    }
  });

  return { frozen, skipped };
}
