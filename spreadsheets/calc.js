// A closed formula language. Cell formula text is never executed as JavaScript.
const MAX_TOKENS = 384;
const MAX_DEPTH = 64;
const MAX_RANGE = 20000;
const MAX_WORK = 100000;
const NUMBER = /^(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?$/i;
const FUNCTIONS = new Set(["SUM", "AVERAGE", "MIN", "MAX", "IF"]);
const PRECEDENCE = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "+": 2, "-": 2, "*": 3, "/": 3, "^": 4
};

class FormulaError extends Error {}
function fail(message) { throw new FormulaError(message); }

function coordinates(text) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/.exec(text);
  if (!match) return null;
  let col = 0;
  for (const letter of match[1]) col = col * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (col > 16384 || row > 1048576) fail("#REF!");
  return { col, row, address: match[1] + match[2] };
}

function columnName(col) {
  let s = "";
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

function tokenize(formula) {
  if (typeof formula !== "string" || formula.length > 4096) fail("#LIMIT!");
  const pattern = /\s+|"(?:[^"]|"")*"|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[A-Za-z_$][A-Za-z0-9_$]*|<>|<=|>=|[()+*/^,:=<>-]/gy;
  const tokens = [];
  let pos = 0;
  while (pos < formula.length) {
    pattern.lastIndex = pos;
    const match = pattern.exec(formula);
    if (!match) fail("#NAME?");
    pos = pattern.lastIndex;
    if (match[0].trim()) tokens.push(match[0]);
    if (tokens.length > MAX_TOKENS) fail("#LIMIT!");
  }
  return tokens;
}

function parse(formula) {
  const tokens = tokenize(formula);
  let pos = 0;
  function take(token) {
    if (tokens[pos] !== token) return false;
    pos++;
    return true;
  }
  function expect(token) { if (!take(token)) fail("#VALUE!"); }
  function reference() {
    const ref = coordinates((tokens[pos++] || "").toUpperCase());
    if (!ref) fail("#REF!");
    return ref;
  }
  function primary(depth) {
    if (depth > MAX_DEPTH) fail("#LIMIT!");
    if (take("(")) {
      const node = expression(1, depth + 1);
      expect(")");
      return node;
    }
    const token = tokens[pos++];
    if (!token) fail("#VALUE!");
    if (token === "-" || token === "+") {
      const arg = primary(depth + 1);
      if (arg.op === "range") fail("#VALUE!");
      return { op: "sign", sign: token, arg };
    }
    if (NUMBER.test(token)) {
      const value = Number(token);
      if (!Number.isFinite(value)) fail("#NUM!");
      return { op: "value", value };
    }
    if (token.startsWith('"')) return { op: "value", value: token.slice(1,-1).replace(/""/g, '"') };
    const name = token.toUpperCase();
    if (take("(")) {
      if (!FUNCTIONS.has(name)) fail("#NAME?");
      const args = [];
      do {
        const node = expression(1, depth + 1);
        if (name === "IF" && node.op === "range") fail("#VALUE!");
        args.push(node);
      } while (take(","));
      expect(")");
      if (name === "IF" && (args.length < 2 || args.length > 3)) fail("#VALUE!");
      return { op: "call", name, args };
    }
    if (name === "TRUE" || name === "FALSE") return { op: "value", value: name === "TRUE" };
    const ref = coordinates(name);
    if (!ref) fail("#NAME?");
    if (!take(":")) return { op: "ref", address: ref.address };
    const end = reference();
    const area = (Math.abs(ref.col - end.col) + 1) * (Math.abs(ref.row - end.row) + 1);
    if (area > MAX_RANGE) fail("#LIMIT!");
    return { op: "range", start: ref, end };
  }
  function expression(minPrecedence, depth) {
    let node = primary(depth);
    while (Object.hasOwn(PRECEDENCE, tokens[pos]) && PRECEDENCE[tokens[pos]] >= minPrecedence) {
      const op = tokens[pos++];
      const right = expression(PRECEDENCE[op] + 1, depth + 1);
      if (node.op === "range" || right.op === "range") fail("#VALUE!");
      node = { op, left: node, right };
    }
    return node;
  }
  const root = expression(1, 0);
  if (pos !== tokens.length || root.op === "range") fail("#VALUE!");
  return root;
}

export function isSupported(formula) {
  try { parse(formula); return true; }
  catch (error) { if (error instanceof FormulaError) return false; throw error; }
}

function finite(value) {
  if (!Number.isFinite(value)) fail("#NUM!");
  return value;
}
function numeric(value) {
  if (value === null) return 0;
  if (typeof value === "string" && !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?$/i.test(value.trim())) fail("#VALUE!");
  return finite(Number(value));
}
function truth(value) {
  if (typeof value !== "string") return Boolean(value);
  if (value.toUpperCase() === "TRUE") return true;
  if (value.toUpperCase() === "FALSE") return false;
  fail("#VALUE!");
}
function compare(a, b) {
  const blankAs = other => typeof other === "string" ? "" : typeof other === "boolean" ? false : 0;
  if (a === null) a = blankAs(b);
  if (b === null) b = blankAs(a);
  if (typeof a !== typeof b) {
    const order = { number: 0, string: 1, boolean: 2 };
    return order[typeof a] - order[typeof b];
  }
  if (typeof a === "string") { a = a.toUpperCase(); b = b.toUpperCase(); }
  return a < b ? -1 : a > b ? 1 : 0;
}
function binary(op, a, b) {
  if (PRECEDENCE[op] === 1) {
    const result = compare(a, b);
    switch (op) {
      case "=": return result === 0;
      case "<>": return result !== 0;
      case "<": return result < 0;
      case "<=": return result <= 0;
      case ">": return result > 0;
      case ">=": return result >= 0;
    }
  }
  const left = numeric(a), right = numeric(b);
  switch (op) {
    case "+": return finite(left + right);
    case "-": return finite(left - right);
    case "*": return finite(left * right);
    case "/": if (right === 0) fail("#DIV/0!"); return finite(left / right);
    case "^": return finite(left ** right);
  }
}

// Make a new reader after changing any input. Memoized values only live on this reader.
export function makeCalculator(book) {
  const caches = book.sheets.map(() => new Map());
  const parsed = new Map();
  return function calculate(sheetIndex, address) {
    const sheet = book.sheets[sheetIndex];
    if (!sheet) return "#REF!";
    const cache = caches[sheetIndex];
    const visiting = new Set();
    let work = 0;
    function step() { if (++work > MAX_WORK) fail("#LIMIT!"); }
    function read(key) {
      step();
      if (cache.has(key)) return cache.get(key);
      const cell = sheet.cells[key];
      if (!cell) return null;
      // The editor keeps an empty styled cell in the model to preserve its format.
      if (typeof cell.f !== "string") return cell.v === "" ? null : cell.v ?? null;
      if (visiting.has(key)) fail("#CYCLE!");
      if (visiting.size >= MAX_DEPTH) fail("#LIMIT!");
      visiting.add(key);
      try {
        if (!parsed.has(cell.f)) parsed.set(cell.f, parse(cell.f));
        const value = evaluate(parsed.get(cell.f)) ?? 0;
        cache.set(key, value);
        return value;
      } finally { visiting.delete(key); }
    }
    function aggregate(node) {
      let sum = 0, count = 0, min = Infinity, max = -Infinity;
      function include(value, fromCell) {
        if (value === null || (fromCell && typeof value !== "number")) return;
        const n = numeric(value);
        sum += n;
        count++;
        min = Math.min(min, n);
        max = Math.max(max, n);
      }
      for (const arg of node.args) {
        if (arg.op !== "range") { include(evaluate(arg), arg.op === "ref"); continue; }
        const { start, end } = arg;
        for (let col = Math.min(start.col,end.col); col <= Math.max(start.col,end.col); col++) {
          const label = columnName(col);
          for (let row = Math.min(start.row,end.row); row <= Math.max(start.row,end.row); row++) {
            include(read(label + row), true);
          }
        }
      }
      switch (node.name) {
        case "SUM": return finite(sum);
        case "AVERAGE": if (!count) fail("#DIV/0!"); return finite(sum / count);
        case "MIN": return count ? min : 0;
        case "MAX": return count ? max : 0;
      }
    }
    function evaluate(node) {
      step();
      switch (node.op) {
        case "value": return node.value;
        case "ref": return read(node.address);
        case "sign": return numeric(evaluate(node.arg)) * (node.sign === "-" ? -1 : 1);
        case "call":
          if (node.name !== "IF") return aggregate(node);
          if (truth(evaluate(node.args[0]))) return evaluate(node.args[1]);
          return node.args.length === 3 ? evaluate(node.args[2]) : false;
        default: return binary(node.op, evaluate(node.left), evaluate(node.right));
      }
    }
    try {
      const ref = coordinates(String(address).toUpperCase());
      if (!ref) return "#REF!";
      return read(ref.address) ?? "";
    } catch (error) {
      if (error instanceof FormulaError) return error.message;
      if (error instanceof RangeError) return "#LIMIT!";
      throw error;
    }
  };
}
