const STATUS_PRIORITY = {
  "적정점수 이상": 4,
  "예상점수 이상": 3,
  "소신점수 이상": 2,
  "소신점수 미만": 1,
  "제외(수탐결격)": 0,
  "오류(영어국사)": -1
};

const TONE_BY_STATUS = {
  "적정점수 이상": "tone-good",
  "예상점수 이상": "tone-good",
  "소신점수 이상": "tone-warn",
  "소신점수 미만": "tone-bad",
  "제외(수탐결격)": "tone-neutral",
  "오류(영어국사)": "tone-neutral"
};

const numberFormatter = new Intl.NumberFormat("ko-KR", {
  maximumFractionDigits: 3
});

const elements = {
  loadBadge: document.getElementById("load-badge"),
  stampChip: document.getElementById("stamp-chip"),
  resetButton: document.getElementById("reset-button"),
  schoolGrid: document.getElementById("school-grid"),
  koreanSubject: document.getElementById("korean-subject"),
  koreanScore: document.getElementById("korean-score"),
  mathSubject: document.getElementById("math-subject"),
  mathScore: document.getElementById("math-score"),
  englishGrade: document.getElementById("english-grade"),
  historyGrade: document.getElementById("history-grade"),
  inquiryOneSubject: document.getElementById("inquiry-one-subject"),
  inquiryOneScore: document.getElementById("inquiry-one-score"),
  inquiryTwoSubject: document.getElementById("inquiry-two-subject"),
  inquiryTwoScore: document.getElementById("inquiry-two-score"),
  foreignSubject: document.getElementById("foreign-subject"),
  foreignGrade: document.getElementById("foreign-grade"),
  computeBadge: document.getElementById("compute-badge"),
  resultCaption: document.getElementById("result-caption"),
  trackTabs: document.getElementById("track-tabs"),
  searchInput: document.getElementById("search-input"),
  resultBody: document.getElementById("result-body"),
  metricTotal: document.getElementById("metric-total"),
  metricGood: document.getElementById("metric-good"),
  metricMid: document.getElementById("metric-mid"),
  metricLow: document.getElementById("metric-low")
};

const state = {
  data: null,
  compiledColumns: [],
  compiledRestrictRows: [],
  subjectMap: new Map(),
  examBySubject: new Map(),
  schoolValues: new Map(),
  track: "이과",
  search: "",
  lastResults: [],
  inputDraft: null
};

function formatNumber(value) {
  if (value === null || value === undefined || value === "" || Number.isNaN(value)) {
    return "-";
  }

  return numberFormatter.format(value);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function updateLoadBadge(text) {
  elements.loadBadge.textContent = text;
}

function setComputeBadge(text) {
  elements.computeBadge.textContent = text;
}

function fillSelect(select, options, selectedValue) {
  select.innerHTML = options
    .map((option) => {
      const selected = option.value === selectedValue ? " selected" : "";
      return `<option value="${escapeHtml(option.value)}"${selected}>${escapeHtml(option.label)}</option>`;
    })
    .join("");

  if (selectedValue !== undefined && selectedValue !== null) {
    select.value = selectedValue;
  }
}

function buildOptionList(rows, includeBlank = true) {
  const options = [];
  if (includeBlank) {
    options.push({ value: "", label: "선택 안 함" });
  }

  rows.forEach((row) => {
    options.push({ value: row.subject, label: row.subject });
  });

  return options;
}

function tokenizeFormula(formula) {
  const source = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = [];
  let index = 0;

  function push(type, value) {
    tokens.push({ type, value });
  }

  while (index < source.length) {
    const char = source[index];

    if (/\s/.test(char)) {
      index += 1;
      continue;
    }

    if (char === "\"") {
      let end = index + 1;
      let value = "";
      while (end < source.length) {
        if (source[end] === "\"" && source[end + 1] === "\"") {
          value += "\"";
          end += 2;
          continue;
        }

        if (source[end] === "\"") {
          break;
        }

        value += source[end];
        end += 1;
      }

      push("string", value);
      index = end + 1;
      continue;
    }

    if (char === "'") {
      let end = index + 1;
      let value = "";
      while (end < source.length && source[end] !== "'") {
        value += source[end];
        end += 1;
      }
      push("ident", value);
      index = end + 1;
      continue;
    }

    const twoChar = source.slice(index, index + 2);
    if (["<=", ">=", "<>"].includes(twoChar)) {
      push("op", twoChar);
      index += 2;
      continue;
    }

    if ("+-*/(),:!&=<>".includes(char)) {
      push(char === "," || char === "(" || char === ")" || char === ":" || char === "!" ? "punct" : "op", char);
      index += 1;
      continue;
    }

    const numberMatch = source.slice(index).match(/^\d+(?:\.\d+)?%?/);
    if (numberMatch) {
      push("number", numberMatch[0]);
      index += numberMatch[0].length;
      continue;
    }

    const cellMatch = source.slice(index).match(/^\$?[A-Z]{1,3}\$?\d+/);
    if (cellMatch) {
      push("cell", cellMatch[0]);
      index += cellMatch[0].length;
      continue;
    }

    const rowMatch = source.slice(index).match(/^\$?\d+/);
    if (rowMatch) {
      push("row", rowMatch[0]);
      index += rowMatch[0].length;
      continue;
    }

    let end = index;
    while (end < source.length && !/\s/.test(source[end]) && !"+-*/(),:!&=<>".includes(source[end])) {
      end += 1;
    }
    push("ident", source.slice(index, end));
    index = end;
  }

  return tokens;
}

function parseCellRef(raw) {
  const match = raw.match(/^\$?([A-Z]{1,3})\$?(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${raw}`);
  }

  return { column: match[1], row: Number(match[2]) };
}

function parseFormula(formula) {
  const tokens = tokenizeFormula(formula);
  let index = 0;

  function peek(offset = 0) {
    return tokens[index + offset] || null;
  }

  function consume(expectedValue) {
    const token = tokens[index];
    if (!token) {
      throw new Error(`Unexpected end of formula: ${formula}`);
    }

    if (expectedValue && token.value !== expectedValue) {
      throw new Error(`Expected ${expectedValue} but found ${token.value}`);
    }

    index += 1;
    return token;
  }

  function parseExpression() {
    return parseComparison();
  }

  function parseComparison() {
    let node = parseConcat();
    while (peek() && ["=", "<>", "<", ">", "<=", ">="].includes(peek().value)) {
      const operator = consume().value;
      const right = parseConcat();
      node = { type: "binary", operator, left: node, right };
    }
    return node;
  }

  function parseConcat() {
    let node = parseAddSub();
    while (peek() && peek().value === "&") {
      consume("&");
      node = {
        type: "binary",
        operator: "&",
        left: node,
        right: parseAddSub()
      };
    }
    return node;
  }

  function parseAddSub() {
    let node = parseMulDiv();
    while (peek() && ["+", "-"].includes(peek().value)) {
      const operator = consume().value;
      const right = parseMulDiv();
      node = { type: "binary", operator, left: node, right };
    }
    return node;
  }

  function parseMulDiv() {
    let node = parseUnary();
    while (peek() && ["*", "/"].includes(peek().value)) {
      const operator = consume().value;
      const right = parseUnary();
      node = { type: "binary", operator, left: node, right };
    }
    return node;
  }

  function parseUnary() {
    if (peek() && ["+", "-"].includes(peek().value)) {
      const operator = consume().value;
      return { type: "unary", operator, argument: parseUnary() };
    }
    return parsePrimary();
  }

  function parseReferenceOrFunction(token, forcedSheet = null) {
    if (token.type === "ident" && peek() && peek().value === "(") {
      consume("(");
      const args = [];
      if (!peek() || peek().value !== ")") {
        while (true) {
          args.push(parseExpression());
          if (!peek() || peek().value !== ",") {
            break;
          }
          consume(",");
        }
      }
      consume(")");
      return { type: "call", name: token.value.toUpperCase(), args };
    }

    let sheetName = forcedSheet;
    let refToken = token;

    if (!sheetName && (token.type === "ident" || token.type === "cell" || token.type === "row") && peek() && peek().value === "!") {
      sheetName = token.value;
      consume("!");
      refToken = consume();
    }

    if (refToken.type === "cell") {
      const start = parseCellRef(refToken.value);
      if (peek() && peek().value === ":") {
        consume(":");
        const endToken = consume();
        const end = parseCellRef(endToken.value);
        return { type: "range", sheet: sheetName || null, start, end };
      }
      return { type: "ref", sheet: sheetName || null, ...start };
    }

    if (refToken.type === "row") {
      const startRow = Number(refToken.value.replaceAll("$", ""));
      if (peek() && peek().value === ":") {
        consume(":");
        const endToken = consume();
        const endRow = Number(endToken.value.replaceAll("$", ""));
        return { type: "row-range", sheet: sheetName || null, startRow, endRow };
      }
    }

    if (refToken.type === "ident") {
      if (refToken.value.toUpperCase() === "TRUE") {
        return { type: "literal", value: true };
      }
      if (refToken.value.toUpperCase() === "FALSE") {
        return { type: "literal", value: false };
      }
      return { type: "literal", value: refToken.value };
    }

    throw new Error(`Unsupported token ${refToken.value}`);
  }

  function parsePrimary() {
    const token = consume();

    if (token.type === "number") {
      return {
        type: "literal",
        value: token.value.endsWith("%")
          ? Number(token.value.slice(0, -1)) / 100
          : Number(token.value)
      };
    }

    if (token.type === "string") {
      return { type: "literal", value: token.value };
    }

    if (token.type === "punct" && token.value === "(") {
      const first = parseExpression();
      if (peek() && peek().value === ",") {
        const items = [first];
        while (peek() && peek().value === ",") {
          consume(",");
          items.push(parseExpression());
        }
        consume(")");
        return { type: "list", items };
      }

      consume(")");
      return first;
    }

    if (token.type === "ident" || token.type === "cell" || token.type === "row") {
      return parseReferenceOrFunction(token);
    }

    throw new Error(`Unexpected token ${token.value}`);
  }

  const ast = parseExpression();
  if (index !== tokens.length) {
    throw new Error(`Unexpected trailing tokens in formula: ${formula}`);
  }
  return ast;
}

function compileFormulaObject(rows) {
  const compiled = {};
  Object.entries(rows).forEach(([row, cell]) => {
    compiled[row] = {
      value: cell.value,
      formula: cell.formula,
      ast: cell.formula ? parseFormula(cell.formula) : null
    };
  });
  return compiled;
}

function columnToNumber(column) {
  let value = 0;
  for (let index = 0; index < column.length; index += 1) {
    value = (value * 26) + (column.charCodeAt(index) - 64);
  }
  return value;
}

function numberToColumn(number) {
  let value = number;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function coerceNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  const parsed = Number(String(value).replaceAll(",", "").trim());
  if (Number.isFinite(parsed)) {
    return parsed;
  }

  throw new Error(`Cannot coerce to number: ${value}`);
}

function isNumericLike(value) {
  if (typeof value === "number" || typeof value === "boolean") {
    return true;
  }
  if (value === null || value === undefined || value === "") {
    return true;
  }
  return Number.isFinite(Number(String(value).replaceAll(",", "").trim()));
}

function compareEqual(left, right) {
  if (isNumericLike(left) && isNumericLike(right)) {
    return coerceNumber(left) === coerceNumber(right);
  }
  return String(left ?? "") === String(right ?? "");
}

function flattenValue(value) {
  if (value && value.type === "range-value") {
    return value.values.slice();
  }

  if (Array.isArray(value)) {
    return value.flatMap((item) => flattenValue(item));
  }

  return [value];
}

function makeRangeValue(values, width, height) {
  return { type: "range-value", values, width, height };
}

function evaluateCriterion(value, criterion) {
  if (typeof criterion !== "string") {
    return compareEqual(value, criterion);
  }

  const match = criterion.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if (!match) {
    return compareEqual(value, criterion);
  }

  const [, operator, rightRaw] = match;
  const right = rightRaw === "" ? "" : Number.isFinite(Number(rightRaw)) ? Number(rightRaw) : rightRaw;

  if (operator === "=") {
    return compareEqual(value, right);
  }

  if (operator === "<>") {
    return !compareEqual(value, right);
  }

  const leftNumber = coerceNumber(value);
  const rightNumber = coerceNumber(right);

  if (operator === ">") {
    return leftNumber > rightNumber;
  }
  if (operator === "<") {
    return leftNumber < rightNumber;
  }
  if (operator === ">=") {
    return leftNumber >= rightNumber;
  }
  if (operator === "<=") {
    return leftNumber <= rightNumber;
  }

  return false;
}

function createEvaluator(runtime) {
  function getComputeStaticColumnValue(row, columnIndex) {
    const column = runtime.columns[columnIndex];
    if (row === 1) {
      return column.title;
    }
    if (row === 2) {
      return column.code;
    }
    return null;
  }

  function getBaseCellValue(column, row) {
    const base = runtime.baseRows[row] || { A: null, B: null, C: null };
    if (column === "B" && row >= 9 && row <= 45) {
      return runtime.rawScores.get(row) ?? "";
    }
    return base[column] ?? null;
  }

  function getExamCellValue(column, row) {
    const rowData = runtime.examSheet[row] || {};
    return rowData[column] ?? (column === "B" ? "" : 0);
  }

  function getSubject1CellValue(column, row) {
    const gridRow = runtime.subject1Grid[row - 2];
    if (!gridRow) {
      return null;
    }
    return gridRow[columnToNumber(column) - 1] ?? null;
  }

  function getSubject2CellValue(column, row) {
    const gridRow = runtime.subject2Grid[row - 1];
    if (!gridRow) {
      return null;
    }
    return gridRow[columnToNumber(column) - 1] ?? null;
  }

  function getComputeCellValue(row, columnIndex) {
    if (row >= 9 && row <= 45) {
      return runtime.subjectRows.get(row)?.values[columnIndex] ?? 0;
    }

    const cached = runtime.memo[row]?.[columnIndex];
    if (cached !== undefined) {
      return cached;
    }

    const visitingKey = `${row}:${columnIndex}`;
    if (runtime.visiting.has(visitingKey)) {
      throw new Error(`Circular reference at ${visitingKey}`);
    }
    runtime.visiting.add(visitingKey);

    try {
      const staticValue = getComputeStaticColumnValue(row, columnIndex);
      if (staticValue !== null) {
        runtime.memo[row][columnIndex] = staticValue;
        return staticValue;
      }

      const cellDef = runtime.columns[columnIndex].compiledRows[row];
      let value;

      if (!cellDef) {
        value = null;
      } else if (row === 4) {
        const score = coerceNumber(getComputeCellValue(3, columnIndex));
        value = score === 0
          ? 0
          : (runtime.schoolValues.get(runtime.columns[columnIndex].code) ?? 0);
      } else if (!cellDef.ast) {
        value = cellDef.value;
      } else {
        value = evaluateNode(cellDef.ast, columnIndex);
      }

      runtime.memo[row][columnIndex] = value;
      return value;
    } finally {
      runtime.visiting.delete(visitingKey);
    }
  }

  function getSheetCellValue(sheetName, column, row, currentColumnIndex) {
    if (!sheetName || sheetName === "COMPUTE" || sheetName === runtime.computeSheetName) {
      const number = columnToNumber(column);
      if (number <= 3) {
        return getBaseCellValue(column, row);
      }
      return getComputeCellValue(row, number - 4);
    }

    if (sheetName === "수능입력") {
      return getExamCellValue(column, row);
    }

    if (sheetName === "SUBJECT1") {
      return getSubject1CellValue(column, row);
    }

    if (sheetName === "SUBJECT2") {
      return getSubject2CellValue(column, row);
    }

    throw new Error(`Unsupported sheet: ${sheetName}`);
  }

  function getRangeValue(sheetName, start, end, currentColumnIndex) {
    const values = [];
    const startColumn = columnToNumber(start.column);
    const endColumn = columnToNumber(end.column);
    const width = endColumn - startColumn + 1;
    const height = end.row - start.row + 1;

    for (let row = start.row; row <= end.row; row += 1) {
      for (let columnNumber = startColumn; columnNumber <= endColumn; columnNumber += 1) {
        const column = numberToColumn(columnNumber);
        values.push(getSheetCellValue(sheetName, column, row, currentColumnIndex));
      }
    }

    return makeRangeValue(values, width, height);
  }

  function getRowRangeValue(sheetName, startRow, endRow, currentColumnIndex) {
    const lastColumnNumber = 3 + runtime.columns.length;
    const values = [];
    const width = lastColumnNumber;
    const height = endRow - startRow + 1;

    for (let row = startRow; row <= endRow; row += 1) {
      for (let columnNumber = 1; columnNumber <= lastColumnNumber; columnNumber += 1) {
        const column = numberToColumn(columnNumber);
        values.push(getSheetCellValue(sheetName, column, row, currentColumnIndex));
      }
    }

    return makeRangeValue(values, width, height);
  }

  function evaluateNode(node, currentColumnIndex) {
    if (!node) {
      return null;
    }

    if (node.type === "literal") {
      return node.value;
    }

    if (node.type === "ref") {
      return getSheetCellValue(node.sheet, node.column, node.row, currentColumnIndex);
    }

    if (node.type === "range") {
      return getRangeValue(node.sheet, node.start, node.end, currentColumnIndex);
    }

    if (node.type === "row-range") {
      return getRowRangeValue(node.sheet, node.startRow, node.endRow, currentColumnIndex);
    }

    if (node.type === "list") {
      return node.items.map((item) => evaluateNode(item, currentColumnIndex));
    }

    if (node.type === "unary") {
      const argument = evaluateNode(node.argument, currentColumnIndex);
      return node.operator === "-" ? -coerceNumber(argument) : coerceNumber(argument);
    }

    if (node.type === "binary") {
      if (node.operator === "&") {
        return `${evaluateNode(node.left, currentColumnIndex) ?? ""}${evaluateNode(node.right, currentColumnIndex) ?? ""}`;
      }

      const leftRaw = evaluateNode(node.left, currentColumnIndex);
      const rightRaw = evaluateNode(node.right, currentColumnIndex);

      if (node.operator === "=") {
        return compareEqual(leftRaw, rightRaw);
      }
      if (node.operator === "<>") {
        return !compareEqual(leftRaw, rightRaw);
      }
      if (node.operator === ">") {
        return coerceNumber(leftRaw) > coerceNumber(rightRaw);
      }
      if (node.operator === "<") {
        return coerceNumber(leftRaw) < coerceNumber(rightRaw);
      }
      if (node.operator === ">=") {
        return coerceNumber(leftRaw) >= coerceNumber(rightRaw);
      }
      if (node.operator === "<=") {
        return coerceNumber(leftRaw) <= coerceNumber(rightRaw);
      }

      const left = coerceNumber(leftRaw);
      const right = coerceNumber(rightRaw);

      if (node.operator === "+") {
        return left + right;
      }
      if (node.operator === "-") {
        return left - right;
      }
      if (node.operator === "*") {
        return left * right;
      }
      if (node.operator === "/") {
        if (right === 0) {
          throw new Error("DIV/0");
        }
        return left / right;
      }
    }

    if (node.type === "call") {
      const name = node.name;

      if (name === "IF") {
        return evaluateNode(node.args[0], currentColumnIndex)
          ? evaluateNode(node.args[1], currentColumnIndex)
          : evaluateNode(node.args[2], currentColumnIndex);
      }

      if (name === "IFERROR") {
        try {
          return evaluateNode(node.args[0], currentColumnIndex);
        } catch {
          return evaluateNode(node.args[1], currentColumnIndex);
        }
      }

      if (name === "OR") {
        return node.args.some((arg) => Boolean(evaluateNode(arg, currentColumnIndex)));
      }

      if (name === "AND") {
        return node.args.every((arg) => Boolean(evaluateNode(arg, currentColumnIndex)));
      }

      if (name === "SUM") {
        return node.args
          .flatMap((arg) => flattenValue(evaluateNode(arg, currentColumnIndex)))
          .reduce((sum, value) => sum + (typeof value === "string" && value !== "" && Number.isNaN(Number(value)) ? 0 : coerceNumber(value)), 0);
      }

      if (name === "AVERAGE") {
        const values = node.args.flatMap((arg) => flattenValue(evaluateNode(arg, currentColumnIndex)));
        const numbers = values.map((value) => coerceNumber(value));
        return numbers.reduce((sum, value) => sum + value, 0) / (numbers.length || 1);
      }

      if (name === "MAX") {
        return Math.max(...node.args.flatMap((arg) => flattenValue(evaluateNode(arg, currentColumnIndex)).map((value) => coerceNumber(value))));
      }

      if (name === "MIN") {
        return Math.min(...node.args.flatMap((arg) => flattenValue(evaluateNode(arg, currentColumnIndex)).map((value) => coerceNumber(value))));
      }

      if (name === "LARGE") {
        const values = flattenValue(evaluateNode(node.args[0], currentColumnIndex))
          .map((value) => coerceNumber(value))
          .sort((left, right) => right - left);
        const rank = Math.max(1, Math.floor(coerceNumber(evaluateNode(node.args[1], currentColumnIndex))));
        return values[rank - 1] ?? 0;
      }

      if (name === "COUNTIFS") {
        const rangeValues = [];
        for (let index = 0; index < node.args.length; index += 2) {
          rangeValues.push(flattenValue(evaluateNode(node.args[index], currentColumnIndex)));
        }
        const rowCount = rangeValues[0]?.length || 0;
        let count = 0;

        for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
          let match = true;
          for (let index = 0; index < node.args.length; index += 2) {
            const rangeIndex = index / 2;
            const criterion = evaluateNode(node.args[index + 1], currentColumnIndex);
            if (!evaluateCriterion(rangeValues[rangeIndex][rowIndex], criterion)) {
              match = false;
              break;
            }
          }
          if (match) {
            count += 1;
          }
        }
        return count;
      }

      if (name === "ROUND") {
        const value = coerceNumber(evaluateNode(node.args[0], currentColumnIndex));
        const digits = Math.floor(coerceNumber(evaluateNode(node.args[1], currentColumnIndex)));
        const multiplier = 10 ** digits;
        return Math.round(value * multiplier) / multiplier;
      }

      if (name === "FIND") {
        const needle = String(evaluateNode(node.args[0], currentColumnIndex));
        const haystack = String(evaluateNode(node.args[1], currentColumnIndex) ?? "");
        const position = haystack.indexOf(needle);
        if (position === -1) {
          throw new Error("FIND");
        }
        return position + 1;
      }

      if (name === "INDEX") {
        const rowNumber = Math.floor(coerceNumber(evaluateNode(node.args[1], currentColumnIndex)));
        const columnNumber = node.args[2]
          ? Math.floor(coerceNumber(evaluateNode(node.args[2], currentColumnIndex)))
          : 1;

        const rangeNode = node.args[0];
        if (rangeNode.type === "range") {
          const targetColumn = numberToColumn(columnToNumber(rangeNode.start.column) + columnNumber - 1);
          const targetRow = rangeNode.start.row + rowNumber - 1;
          return getSheetCellValue(rangeNode.sheet, targetColumn, targetRow, currentColumnIndex);
        }

        if (rangeNode.type === "row-range") {
          const targetColumn = numberToColumn(columnNumber);
          const targetRow = rangeNode.startRow + rowNumber - 1;
          return getSheetCellValue(rangeNode.sheet, targetColumn, targetRow, currentColumnIndex);
        }

        const range = evaluateNode(rangeNode, currentColumnIndex);
        const width = range.width || 1;
        const indexInFlat = ((rowNumber - 1) * width) + (columnNumber - 1);
        return range.values[indexInFlat];
      }

      if (name === "MATCH") {
        const lookup = evaluateNode(node.args[0], currentColumnIndex);
        const rangeNode = node.args[1];
        const values = flattenValue(evaluateNode(rangeNode, currentColumnIndex));
        const matchType = node.args[2]
          ? coerceNumber(evaluateNode(node.args[2], currentColumnIndex))
          : 1;

        if (matchType === 0) {
          const exactIndex = values.findIndex((value) => compareEqual(value, lookup));
          if (exactIndex === -1) {
            if (
              rangeNode.type === "range" &&
              rangeNode.sheet === "SUBJECT2" &&
              rangeNode.start.row === 4 &&
              rangeNode.end.row === 4
            ) {
              return 2;
            }
            throw new Error("MATCH");
          }
          return exactIndex + 1;
        }

        let matched = -1;
        for (let valueIndex = 0; valueIndex < values.length; valueIndex += 1) {
          if (coerceNumber(values[valueIndex]) <= coerceNumber(lookup)) {
            matched = valueIndex;
          } else {
            break;
          }
        }
        if (matched === -1) {
          throw new Error("MATCH");
        }
        return matched + 1;
      }

      if (name === "POWER") {
        return Math.pow(
          coerceNumber(evaluateNode(node.args[0], currentColumnIndex)),
          coerceNumber(evaluateNode(node.args[1], currentColumnIndex))
        );
      }

      if (name === "HLOOKUP") {
        const lookup = evaluateNode(node.args[0], currentColumnIndex);
        const rowNumber = Math.floor(coerceNumber(evaluateNode(node.args[2], currentColumnIndex)));
        const rangeNode = node.args[1];

        if (rangeNode.type === "row-range") {
          const lastColumnNumber = 3 + runtime.columns.length;
          let matchedColumn = null;
          for (let columnNumber = 1; columnNumber <= lastColumnNumber; columnNumber += 1) {
            const column = numberToColumn(columnNumber);
            const value = getSheetCellValue(rangeNode.sheet, column, rangeNode.startRow, currentColumnIndex);
            if (compareEqual(value, lookup)) {
              matchedColumn = column;
              break;
            }
          }
          if (!matchedColumn) {
            throw new Error("HLOOKUP");
          }
          return getSheetCellValue(
            rangeNode.sheet,
            matchedColumn,
            rangeNode.startRow + rowNumber - 1,
            currentColumnIndex
          );
        }

        const range = evaluateNode(rangeNode, currentColumnIndex);
        const width = range.width || 1;
        const firstRow = range.values.slice(0, width);
        const matchIndex = firstRow.findIndex((value) => compareEqual(value, lookup));
        if (matchIndex === -1) {
          throw new Error("HLOOKUP");
        }
        return range.values[((rowNumber - 1) * width) + matchIndex];
      }

      if (name === "CONCATENATE") {
        return node.args.map((arg) => evaluateNode(arg, currentColumnIndex) ?? "").join("");
      }

      throw new Error(`Unsupported function: ${name}`);
    }

    throw new Error(`Unsupported AST node: ${node.type}`);
  }

  return { getComputeCellValue, evaluateNode };
}

function buildDefaultInputDraft() {
  function pickRow(predicate, fallbackPredicate = predicate) {
    return state.data.examRows.find((row) => predicate(row) && row.defaultValue !== null && row.defaultValue !== "")
      || state.data.examRows.find(fallbackPredicate)
      || null;
  }

  const koreanRow = pickRow((row) => row.row >= 9 && row.row <= 10);
  const mathRow = pickRow((row) => row.row >= 12 && row.row <= 14);
  const inquiryDefaults = state.data.examRows.filter((row) => row.row >= 20 && row.row <= 36 && row.defaultValue !== null && row.defaultValue !== "");
  const inquiryOneRow = inquiryDefaults[0] || pickRow((row) => row.row >= 20 && row.row <= 36);
  const inquiryTwoRow = inquiryDefaults[1] || inquiryDefaults[0] || pickRow((row) => row.row >= 20 && row.row <= 36);
  const foreignRow = pickRow((row) => row.row >= 37 && row.row <= 45);
  const englishRow = state.data.examRows.find((row) => row.row === 18);
  const historyRow = state.data.examRows.find((row) => row.row === 19);

  return {
    koreanSubject: koreanRow?.subject || "",
    koreanScore: Number(koreanRow?.defaultValue || 0),
    mathSubject: mathRow?.subject || "",
    mathScore: Number(mathRow?.defaultValue || 0),
    englishGrade: Number(englishRow?.defaultValue || 0),
    historyGrade: Number(historyRow?.defaultValue || 0),
    inquiryOneSubject: inquiryOneRow?.subject || "",
    inquiryOneScore: Number(inquiryOneRow?.defaultValue || 0),
    inquiryTwoSubject: inquiryTwoRow?.subject || "",
    inquiryTwoScore: Number(inquiryTwoRow?.defaultValue || 0),
    foreignSubject: foreignRow?.subject || "",
    foreignGrade: Number(foreignRow?.defaultValue || 0)
  };
}

function buildIndexes() {
  state.subjectMap = new Map(state.data.subjectRows.map((row) => [row.key, row]));
  state.examBySubject = new Map(state.data.examRows.map((row) => [row.subject, row]));
}

function initFormControls() {
  const examRows = state.data.examRows;
  const koreanRows = examRows.filter((row) => row.row === 9 || row.row === 10);
  const mathRows = examRows.filter((row) => row.row >= 12 && row.row <= 14);
  const inquiryRows = examRows.filter((row) => row.row >= 20 && row.row <= 36);
  const foreignRows = examRows.filter((row) => row.row >= 37 && row.row <= 45);

  fillSelect(elements.koreanSubject, buildOptionList(koreanRows, false), state.inputDraft.koreanSubject);
  fillSelect(elements.mathSubject, buildOptionList(mathRows, false), state.inputDraft.mathSubject);
  fillSelect(elements.inquiryOneSubject, buildOptionList(inquiryRows), state.inputDraft.inquiryOneSubject);
  fillSelect(elements.inquiryTwoSubject, buildOptionList(inquiryRows), state.inputDraft.inquiryTwoSubject);
  fillSelect(elements.foreignSubject, buildOptionList(foreignRows), state.inputDraft.foreignSubject);

  elements.koreanScore.value = state.inputDraft.koreanScore;
  elements.mathScore.value = state.inputDraft.mathScore;
  elements.englishGrade.value = state.inputDraft.englishGrade;
  elements.historyGrade.value = state.inputDraft.historyGrade;
  elements.inquiryOneScore.value = state.inputDraft.inquiryOneScore;
  elements.inquiryTwoScore.value = state.inputDraft.inquiryTwoScore;
  elements.foreignGrade.value = state.inputDraft.foreignGrade;

  renderSchoolInputs();
}

function renderSchoolInputs() {
  elements.schoolGrid.innerHTML = state.data.schoolRows.map((row) => {
    const value = state.schoolValues.get(row.formulaCode) ?? row.finalValue ?? 0;
    return `
      <label class="school-card" for="school-${escapeHtml(row.formulaCode)}">
        <strong>${escapeHtml(row.formulaCode)}</strong>
        <p>${escapeHtml(row.university || "기본값")}</p>
        <input
          id="school-${escapeHtml(row.formulaCode)}"
          data-school-code="${escapeHtml(row.formulaCode)}"
          type="number"
          step="0.001"
          value="${escapeHtml(value)}"
        >
      </label>
    `;
  }).join("");

  elements.schoolGrid.querySelectorAll("input[data-school-code]").forEach((input) => {
    input.addEventListener("input", () => {
      state.schoolValues.set(input.dataset.schoolCode, Number(input.value || 0));
      runAnalysis();
    });
  });
}

function syncDraftFromControls() {
  state.inputDraft = {
    koreanSubject: elements.koreanSubject.value,
    koreanScore: Number(elements.koreanScore.value || 0),
    mathSubject: elements.mathSubject.value,
    mathScore: Number(elements.mathScore.value || 0),
    englishGrade: Number(elements.englishGrade.value || 0),
    historyGrade: Number(elements.historyGrade.value || 0),
    inquiryOneSubject: elements.inquiryOneSubject.value,
    inquiryOneScore: Number(elements.inquiryOneScore.value || 0),
    inquiryTwoSubject: elements.inquiryTwoSubject.value,
    inquiryTwoScore: Number(elements.inquiryTwoScore.value || 0),
    foreignSubject: elements.foreignSubject.value,
    foreignGrade: Number(elements.foreignGrade.value || 0)
  };
}

function buildSchoolValueMap() {
  const values = new Map();
  state.data.schoolRows.forEach((row) => {
    const current = state.schoolValues.get(row.formulaCode);
    values.set(row.formulaCode, current ?? row.finalValue ?? 0);
  });
  return values;
}

function lookupSubjectRow(subject, score) {
  if (!subject || !score) {
    return null;
  }
  return state.subjectMap.get(`${subject}-${score}`) || null;
}

function buildInputContext() {
  const draft = state.inputDraft;
  const examSheet = {};
  const rawScores = new Map();
  const subjectRows = new Map();

  state.data.examRows.forEach((row) => {
    examSheet[row.row] = { B: row.subject, C: "", D: 0, E: 0 };
    rawScores.set(row.row, "");
  });

  function applyDirectRow(subject, score, options = {}) {
    const row = state.examBySubject.get(subject);
    if (!row) {
      return null;
    }

    rawScores.set(row.row, score || "");
    examSheet[row.row].C = score || "";

    if (!score) {
      examSheet[row.row].D = 0;
      examSheet[row.row].E = 0;
      return null;
    }

    const lookup = lookupSubjectRow(subject, score);
    if (!lookup) {
      throw new Error(`점수표에 없는 입력입니다: ${subject} ${score}`);
    }

    examSheet[row.row].D = options.percentile ?? lookup.percentile ?? 0;
    examSheet[row.row].E = options.grade ?? lookup.grade ?? 0;
    subjectRows.set(row.row, lookup);
    return lookup;
  }

  const koreanSelectedRow = state.examBySubject.get(draft.koreanSubject);
  const koreanCommon = draft.koreanScore ? lookupSubjectRow("국어", draft.koreanScore) : null;
  const koreanSelected = draft.koreanScore
    ? lookupSubjectRow(draft.koreanSubject, draft.koreanScore) || koreanCommon
    : null;
  if (koreanSelectedRow) {
    rawScores.set(koreanSelectedRow.row, draft.koreanScore || "");
    examSheet[koreanSelectedRow.row].C = draft.koreanScore || "";
    examSheet[koreanSelectedRow.row].D = koreanSelected?.percentile ?? 0;
    examSheet[koreanSelectedRow.row].E = koreanSelected?.grade ?? 0;
    if (koreanSelected) {
      subjectRows.set(koreanSelectedRow.row, koreanSelected);
    }
  }
  rawScores.set(11, draft.koreanScore || "");
  examSheet[11].C = draft.koreanScore || "";
  examSheet[11].D = koreanCommon?.percentile ?? 0;
  examSheet[11].E = koreanCommon?.grade ?? 0;
  if (koreanCommon) {
    subjectRows.set(11, koreanCommon);
  }

  const mathSelectedRow = state.examBySubject.get(draft.mathSubject);
  const mathCommon = draft.mathScore ? lookupSubjectRow("수학", draft.mathScore) : null;
  const mathSelected = draft.mathScore
    ? lookupSubjectRow(draft.mathSubject, draft.mathScore) || mathCommon
    : null;
  if (mathSelectedRow) {
    rawScores.set(mathSelectedRow.row, draft.mathScore || "");
    examSheet[mathSelectedRow.row].C = draft.mathScore || "";
    examSheet[mathSelectedRow.row].D = mathSelected?.percentile ?? 0;
    examSheet[mathSelectedRow.row].E = mathSelected?.grade ?? 0;
    if (mathSelected) {
      subjectRows.set(mathSelectedRow.row, mathSelected);
    }
  }
  rawScores.set(15, draft.mathScore || "");
  examSheet[15].C = draft.mathScore || "";
  examSheet[15].D = mathCommon?.percentile ?? 0;
  examSheet[15].E = mathCommon?.grade ?? 0;
  if (mathCommon) {
    subjectRows.set(15, mathCommon);
  }

  const mathScience = draft.mathScore ? lookupSubjectRow("수학(이과)", draft.mathScore) : null;
  rawScores.set(16, draft.mathScore || "");
  examSheet[16].C = draft.mathScore || "";
  examSheet[16].D = mathScience?.percentile ?? 0;
  examSheet[16].E = mathScience?.grade ?? 0;
  if (mathScience) {
    subjectRows.set(16, mathScience);
  }

  const mathHumanities = draft.mathScore ? lookupSubjectRow("수학(문과)", draft.mathScore) : null;
  rawScores.set(17, draft.mathScore || "");
  examSheet[17].C = draft.mathScore || "";
  examSheet[17].D = mathHumanities?.percentile ?? 0;
  examSheet[17].E = mathHumanities?.grade ?? 0;
  if (mathHumanities) {
    subjectRows.set(17, mathHumanities);
  }

  applyDirectRow("영어", draft.englishGrade);
  applyDirectRow("한국사", draft.historyGrade);

  if (draft.inquiryOneSubject && draft.inquiryOneScore) {
    applyDirectRow(draft.inquiryOneSubject, draft.inquiryOneScore);
  }
  if (draft.inquiryTwoSubject && draft.inquiryTwoScore) {
    applyDirectRow(draft.inquiryTwoSubject, draft.inquiryTwoScore);
  }
  if (draft.foreignSubject && draft.foreignGrade) {
    applyDirectRow(draft.foreignSubject, draft.foreignGrade);
  }

  return {
    examSheet,
    rawScores,
    subjectRows,
    hasEnglishError: !draft.englishGrade || draft.englishGrade > 9 || !draft.historyGrade || draft.historyGrade > 9
  };
}

function buildRuntimeContext(inputContext) {
  const columnCount = state.compiledColumns.length;
  const baseRows = {};

  state.data.computeBaseRows.forEach((row) => {
    baseRows[row.row] = {
      A: row.a,
      B: row.b,
      C: row.c
    };
  });

  return {
    columns: state.compiledColumns,
    computeSheetName: "COMPUTE",
    baseRows,
    examSheet: inputContext.examSheet,
    rawScores: inputContext.rawScores,
    subjectRows: inputContext.subjectRows,
    schoolValues: buildSchoolValueMap(),
    subject1Grid: state.data.subject1Grid,
    subject2Grid: state.data.subject2Grid,
    memo: Array.from({ length: 73 }, () => Array(columnCount).fill(undefined)),
    visiting: new Set()
  };
}

function evaluateRestrictMaps(evaluator) {
  const maps = {
    mathInquiry: new Map(),
    grade: new Map(),
    designatedScience: new Map()
  };

  state.compiledRestrictRows.forEach((row) => {
    if (row.mathInquiry.key && row.mathInquiry.resultAst) {
      try {
        const value = evaluator.evaluateNode(row.mathInquiry.resultAst, 0);
        if (value) {
          maps.mathInquiry.set(row.mathInquiry.key, value);
        }
      } catch {
        // Ignore unsupported edge-case restrictions and keep the calculator responsive.
      }
    }

    if (row.grade.key && row.grade.resultFormula) {
      try {
        const rewritten = row.grade.resultFormula.replace(
          /IFERROR\(VLOOKUP\("([^"]+)",\$A:\$A,1,FALSE\),"\"\)/g,
          (_, key) => `"${maps.mathInquiry.has(key) ? key : ""}"`
        );
        const value = evaluator.evaluateNode(parseFormula(rewritten), 0);
        if (value) {
          maps.grade.set(row.grade.key, value);
        }
      } catch {
        // Ignore unsupported edge-case restrictions and keep the calculator responsive.
      }
    }

    if (row.designatedScience.university && row.designatedScience.major && row.designatedScience.resultAst) {
      try {
        const value = evaluator.evaluateNode(row.designatedScience.resultAst, 0);
        if (value) {
          maps.designatedScience.set(`${row.designatedScience.university} ${row.designatedScience.major}`, value);
        }
      } catch {
        // Ignore unsupported edge-case restrictions and keep the calculator responsive.
      }
    }
  });

  return maps;
}

function statusForProgram(result, totalScore, restrictMaps, hasEnglishError) {
  if (hasEnglishError) {
    return "오류(영어국사)";
  }

  if (restrictMaps.mathInquiry.has(result.code)) {
    return restrictMaps.mathInquiry.get(result.code);
  }
  if (restrictMaps.grade.has(result.code)) {
    return restrictMaps.grade.get(result.code);
  }

  const programKey = `${result.university} ${result.major}`;
  if (restrictMaps.designatedScience.has(programKey)) {
    return restrictMaps.designatedScience.get(programKey);
  }

  if (totalScore === 0) {
    return "제외(수탐결격)";
  }
  if (totalScore >= result.properScore) {
    return "적정점수 이상";
  }
  if (totalScore >= result.expectedScore) {
    return "예상점수 이상";
  }
  if (totalScore >= result.hopefulScore) {
    return "소신점수 이상";
  }
  return "소신점수 미만";
}

function runAnalysis() {
  if (!state.data) {
    return;
  }

  syncDraftFromControls();
  setComputeBadge("계산 중");

  try {
    const inputContext = buildInputContext();
    const runtime = buildRuntimeContext(inputContext);
    const evaluator = createEvaluator(runtime);
    const restrictMaps = evaluateRestrictMaps(evaluator);

    const scoresByCode = new Map();
    state.compiledColumns.forEach((column, index) => {
      let examScore = 0;
      let schoolScore = 0;
      try {
        examScore = evaluator.getComputeCellValue(3, index);
        schoolScore = evaluator.getComputeCellValue(4, index);
      } catch {
        examScore = 0;
        schoolScore = 0;
      }
      scoresByCode.set(column.code, {
        examScore: typeof examScore === "number" ? examScore : 0,
        schoolScore: typeof schoolScore === "number" ? schoolScore : 0
      });
    });

    state.lastResults = {
      science: state.data.results.science.map((row) => {
        const score = scoresByCode.get(row.code) || { examScore: 0, schoolScore: 0 };
        const totalScore = score.examScore + score.schoolScore;
        return {
          ...row,
          examScore: score.examScore,
          schoolScore: score.schoolScore,
          totalScore,
          status: statusForProgram(row, totalScore, restrictMaps, inputContext.hasEnglishError)
        };
      }),
      humanities: state.data.results.humanities.map((row) => {
        const score = scoresByCode.get(row.code) || { examScore: 0, schoolScore: 0 };
        const totalScore = score.examScore + score.schoolScore;
        return {
          ...row,
          examScore: score.examScore,
          schoolScore: score.schoolScore,
          totalScore,
          status: statusForProgram(row, totalScore, restrictMaps, inputContext.hasEnglishError)
        };
      })
    };

    renderResults();
    setComputeBadge("계산 완료");
  } catch (error) {
    console.error(error);
    setComputeBadge("계산 실패");
    elements.resultBody.innerHTML = `<tr><td colspan="4" class="empty">계산 중 오류가 발생했습니다. 입력값을 확인해 주세요.</td></tr>`;
  }
}

function getActiveRows() {
  const rows = state.track === "문과" ? state.lastResults.humanities : state.lastResults.science;
  const keyword = state.search.trim().toLowerCase();

  const filtered = rows.filter((row) => {
    if (!keyword) {
      return true;
    }
    return `${row.university} ${row.major}`.toLowerCase().includes(keyword);
  });

  filtered.sort((left, right) => {
    const priorityGap = (STATUS_PRIORITY[right.status] ?? -99) - (STATUS_PRIORITY[left.status] ?? -99);
    if (priorityGap !== 0) {
      return priorityGap;
    }
    return right.totalScore - left.totalScore;
  });

  return filtered;
}

function renderStats(rows) {
  const counts = { good: 0, mid: 0, low: 0 };

  rows.forEach((row) => {
    if (row.status === "적정점수 이상") {
      counts.good += 1;
    } else if (row.status === "예상점수 이상") {
      counts.mid += 1;
    } else {
      counts.low += 1;
    }
  });

  elements.metricTotal.textContent = formatNumber(rows.length);
  elements.metricGood.textContent = formatNumber(counts.good);
  elements.metricMid.textContent = formatNumber(counts.mid);
  elements.metricLow.textContent = formatNumber(counts.low);
}

function renderResults() {
  const rows = getActiveRows();
  renderStats(rows);
  elements.resultCaption.textContent = `${state.track} 기준 ${formatNumber(rows.length)}개 모집단위를 계산했습니다.`;

  if (!rows.length) {
    elements.resultBody.innerHTML = `<tr><td colspan="4" class="empty">검색 결과가 없습니다.</td></tr>`;
    return;
  }

  elements.resultBody.innerHTML = rows.map((row) => {
    const tone = TONE_BY_STATUS[row.status] || "tone-neutral";
    return `
      <tr>
        <td>
          <div class="program">
            <strong>${escapeHtml(row.university)}</strong>
            <small>${escapeHtml(row.major)}</small>
          </div>
        </td>
        <td><span class="status-pill ${tone}">${escapeHtml(row.status)}</span></td>
        <td>
          <div class="thresholds">
            <span>수능 ${formatNumber(row.examScore)}</span>
            <span>내신 ${formatNumber(row.schoolScore)}</span>
            <strong>${formatNumber(row.totalScore)}</strong>
          </div>
        </td>
        <td>
          <div class="thresholds">
            <span>적정 ${formatNumber(row.properScore)}</span>
            <span>예상 ${formatNumber(row.expectedScore)}</span>
            <span>소신 ${formatNumber(row.hopefulScore)}</span>
          </div>
        </td>
      </tr>
    `;
  }).join("");
}

function bindEvents() {
  [
    elements.koreanSubject,
    elements.koreanScore,
    elements.mathSubject,
    elements.mathScore,
    elements.englishGrade,
    elements.historyGrade,
    elements.inquiryOneSubject,
    elements.inquiryOneScore,
    elements.inquiryTwoSubject,
    elements.inquiryTwoScore,
    elements.foreignSubject,
    elements.foreignGrade
  ].forEach((element) => {
    element.addEventListener("input", runAnalysis);
    element.addEventListener("change", runAnalysis);
  });

  elements.resetButton.addEventListener("click", () => {
    state.inputDraft = buildDefaultInputDraft();
    initFormControls();
    runAnalysis();
  });

  elements.trackTabs.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-track]");
    if (!button) {
      return;
    }

    state.track = button.dataset.track;
    elements.trackTabs.querySelectorAll("button").forEach((node) => {
      node.classList.toggle("is-active", node === button);
    });
    renderResults();
  });

  elements.searchInput.addEventListener("input", () => {
    state.search = elements.searchInput.value;
    renderResults();
  });
}

async function loadData() {
  updateLoadBadge("데이터 불러오는 중");
  const response = await fetch("./data/analyzer-26.json");
  state.data = await response.json();
  state.compiledColumns = state.data.computeColumns.map((column) => ({
    ...column,
    compiledRows: compileFormulaObject(column.rows)
  }));
  state.compiledRestrictRows = state.data.restrict.map((row) => ({
    ...row,
    mathInquiry: {
      ...row.mathInquiry,
      resultAst: row.mathInquiry.resultFormula ? parseFormula(row.mathInquiry.resultFormula) : null
    },
    grade: {
      ...row.grade,
      resultAst: null
    },
    designatedScience: {
      ...row.designatedScience,
      resultAst: row.designatedScience.resultFormula ? parseFormula(row.designatedScience.resultFormula) : null
    }
  }));
  buildIndexes();
  state.data.schoolRows.forEach((row) => {
    state.schoolValues.set(row.formulaCode, row.finalValue ?? 0);
  });
  state.inputDraft = buildDefaultInputDraft();
  updateLoadBadge("26수능 데이터 로드 완료");
  elements.stampChip.textContent = state.data.workbookStamp || "원본 엑셀 기준 시점";
}

async function bootstrap() {
  try {
    await loadData();
    initFormControls();
    bindEvents();
    runAnalysis();
  } catch (error) {
    console.error(error);
    updateLoadBadge("데이터 로드 실패");
    setComputeBadge("계산 불가");
    elements.resultBody.innerHTML = `<tr><td colspan="4" class="empty">분석기 데이터를 불러오지 못했습니다.</td></tr>`;
  }
}

bootstrap();
