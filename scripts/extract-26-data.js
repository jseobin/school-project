const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const ROOT = path.resolve(__dirname, "..");
const WORKBOOK_PATH = path.join(
  ROOT,
  "..",
  fs.readdirSync(path.join(ROOT, "..")).find((name) => /^202511/.test(name))
);
const OUTPUT_PATH = path.join(ROOT, "data", "analyzer-26.json");

const wb = XLSX.readFile(WORKBOOK_PATH, {
  cellFormula: true,
  cellText: false,
  cellDates: false
});

const SHEETS = {
  exam: wb.SheetNames[1],
  school: wb.SheetNames[2],
  science: wb.SheetNames[3],
  humanities: wb.SheetNames[4],
  compute: wb.SheetNames[5],
  restrict: wb.SheetNames[6],
  subject1: wb.SheetNames[7],
  subject2: wb.SheetNames[8],
  subject3: wb.SheetNames[9]
};

function getCell(sheet, row, col) {
  return sheet[XLSX.utils.encode_cell({ r: row, c: col })] || null;
}

function getCellValue(sheet, row, col) {
  return getCell(sheet, row, col)?.v ?? null;
}

function getText(sheet, row, col) {
  const value = getCellValue(sheet, row, col);
  return value === null || value === undefined ? "" : String(value).trim();
}

function toNumber(value) {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function isBlank(value) {
  return value === null || value === undefined || value === "";
}

function collectExamRows() {
  const sheet = wb.Sheets[SHEETS.exam];
  const rows = [];
  let area = "";

  for (let row = 8; row <= 44; row += 1) {
    const areaText = getText(sheet, row, 0);
    const subject = getText(sheet, row, 1);
    const directValue = getCellValue(sheet, row, 2);

    if (areaText) {
      area = areaText;
    }

    if (!subject) {
      continue;
    }

    rows.push({
      row: row + 1,
      area,
      subject,
      defaultValue: directValue === "" ? null : directValue,
      defaultPercentile: toNumber(getCellValue(sheet, row, 3)),
      defaultGrade: toNumber(getCellValue(sheet, row, 4)),
      defaultCumulative: toNumber(getCellValue(sheet, row, 5))
    });
  }

  return rows;
}

function collectSchoolRows() {
  const sheet = wb.Sheets[SHEETS.school];
  const rows = [];
  let university = "";

  for (let row = 6; row <= 27; row += 1) {
    const universityText = getText(sheet, row, 0);
    const formulaCode = getText(sheet, row, 1);

    if (universityText) {
      university = universityText;
    }

    if (!formulaCode) {
      continue;
    }

    rows.push({
      row: row + 1,
      university,
      formulaCode,
      maxLabel: getCellValue(sheet, row, 2) ?? "",
      directInput: getCellValue(sheet, row, 3) ?? "",
      defaultValue: getCellValue(sheet, row, 4) ?? null,
      finalValue: getCellValue(sheet, row, 5) ?? null
    });
  }

  return rows;
}

function collectComputeColumns() {
  const sheet = wb.Sheets[SHEETS.compute];
  const columns = [];
  const neededRows = new Set([2, 3, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71]);

  for (let col = 3; col <= XLSX.utils.decode_range(sheet["!ref"]).e.c; col += 1) {
    const title = getText(sheet, 0, col);
    const code = getText(sheet, 1, col);

    if (!title && !code) {
      continue;
    }

    const rowMap = {};

    for (let row = 2; row <= 71; row += 1) {
      if (!neededRows.has(row)) {
        continue;
      }

      const cell = getCell(sheet, row, col);
      if (!cell) {
        continue;
      }

      rowMap[row + 1] = {
        value: cell.v ?? null,
        formula: cell.f || null
      };
    }

    columns.push({
      computeColumn: XLSX.utils.encode_col(col),
      index: columns.length,
      title,
      code,
      scorePreview: getCellValue(sheet, 2, col) ?? null,
      schoolPreview: getCellValue(sheet, 3, col) ?? null,
      rows: rowMap
    });
  }

  return columns;
}

function collectComputeBaseRows() {
  const sheet = wb.Sheets[SHEETS.compute];
  const rows = [];

  for (let row = 0; row <= 71; row += 1) {
    rows.push({
      row: row + 1,
      a: getCellValue(sheet, row, 0),
      b: getCellValue(sheet, row, 1),
      c: getCellValue(sheet, row, 2)
    });
  }

  return rows;
}

function collectSubjectRows(columnCount) {
  const sheet = wb.Sheets[SHEETS.subject3];
  const rows = [];

  for (let row = 4; row <= XLSX.utils.decode_range(sheet["!ref"]).e.r; row += 1) {
    const key = getText(sheet, row, 3);
    if (!key) {
      continue;
    }

    const values = [];
    for (let col = 10; col < 10 + columnCount; col += 1) {
      values.push(getCellValue(sheet, row, col) ?? 0);
    }

    rows.push({
      key,
      rawScore: getCellValue(sheet, row, 2),
      percentile: toNumber(getCellValue(sheet, row, 4)),
      percentileWithinSelection: toNumber(getCellValue(sheet, row, 5)),
      grade: toNumber(getCellValue(sheet, row, 6)),
      gradeWithinSelection: toNumber(getCellValue(sheet, row, 7)),
      cumulative: toNumber(getCellValue(sheet, row, 8)),
      values
    });
  }

  return rows;
}

function collectSubject1Meta() {
  const sheet = wb.Sheets[SHEETS.subject1];
  const rows = [];

  for (let row = 4; row <= XLSX.utils.decode_range(sheet["!ref"]).e.r; row += 1) {
    const subject = getText(sheet, row, 0);
    if (!subject) {
      continue;
    }

    rows.push({
      subject,
      maxStandardScore: toNumber(getCellValue(sheet, row, 1)),
      maxPercentile: toNumber(getCellValue(sheet, row, 2)),
      examinees: toNumber(getCellValue(sheet, row, 3))
    });
  }

  return rows;
}

function collectSubject1Grid() {
  const sheet = wb.Sheets[SHEETS.subject1];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows = [];

  for (let row = 1; row <= range.e.r; row += 1) {
    const values = [];
    for (let col = 0; col <= range.e.c; col += 1) {
      values.push(getCellValue(sheet, row, col) ?? null);
    }
    rows.push(values);
  }

  return rows;
}

function collectSubject2Grid() {
  const sheet = wb.Sheets[SHEETS.subject2];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows = [];

  for (let row = 0; row <= range.e.r; row += 1) {
    const values = [];
    for (let col = 0; col <= range.e.c; col += 1) {
      values.push(getCellValue(sheet, row, col) ?? null);
    }
    rows.push(values);
  }

  return rows;
}

function collectRestrictMaps() {
  const sheet = wb.Sheets[SHEETS.restrict];
  const rows = [];

  for (let row = 2; row <= XLSX.utils.decode_range(sheet["!ref"]).e.r; row += 1) {
    rows.push({
      row: row + 1,
      mathInquiry: {
        targetFormula: getCell(sheet, row, 0)?.f || null,
        targetValue: getCellValue(sheet, row, 0),
        key: getText(sheet, row, 1),
        resultFormula: getCell(sheet, row, 2)?.f || null,
        resultValue: getCellValue(sheet, row, 2)
      },
      grade: {
        targetFormula: getCell(sheet, row, 4)?.f || null,
        targetValue: getCellValue(sheet, row, 4),
        key: getText(sheet, row, 5),
        resultFormula: getCell(sheet, row, 6)?.f || null,
        resultValue: getCellValue(sheet, row, 6)
      },
      designatedScience: {
        targetFormula: getCell(sheet, row, 8)?.f || null,
        targetValue: getCellValue(sheet, row, 8),
        university: getText(sheet, row, 9),
        major: getText(sheet, row, 10),
        resultFormula: getCell(sheet, row, 11)?.f || null,
        resultValue: getCellValue(sheet, row, 11)
      }
    });
  }

  return rows;
}

function collectResultRows(sheetName, columnIndex) {
  const sheet = wb.Sheets[sheetName];
  const rows = [];

  for (let row = 5; row <= XLSX.utils.decode_range(sheet["!ref"]).e.r; row += 1) {
    const track = getText(sheet, row, 0);
    const university = getText(sheet, row, 1);
    const major = getText(sheet, row, 2);
    const code = getText(sheet, row, columnIndex);

    if (!track || !university || !major || !code) {
      continue;
    }

    rows.push({
      row: row + 1,
      track,
      university,
      major,
      code,
      properScore: toNumber(getCellValue(sheet, row, 9)),
      expectedScore: toNumber(getCellValue(sheet, row, 10)),
      hopefulScore: toNumber(getCellValue(sheet, row, 11))
    });
  }

  return rows;
}

const computeColumns = collectComputeColumns();
const data = {
  yearKey: "26",
  yearLabel: "26수능",
  workbookStamp: getText(wb.Sheets.INFO, 0, 0),
  examRows: collectExamRows(),
  schoolRows: collectSchoolRows(),
  computeBaseRows: collectComputeBaseRows(),
  computeColumns,
  subjectRows: collectSubjectRows(computeColumns.length),
  subject1Meta: collectSubject1Meta(),
  subject1Grid: collectSubject1Grid(),
  subject2Grid: collectSubject2Grid(),
  restrict: collectRestrictMaps(),
  results: {
    science: collectResultRows(SHEETS.science, 36),
    humanities: collectResultRows(SHEETS.humanities, 36)
  }
};

fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
fs.writeFileSync(OUTPUT_PATH, JSON.stringify(data));

console.log(`Wrote ${OUTPUT_PATH}`);
