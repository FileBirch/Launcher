// A closed formula language. Cell formula text is never executed as JavaScript.
import {address as a1, coords} from './xml.js';
class Unavailable extends Error {}
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
  const pattern = /\s+|#(?:REF!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A|NULL!)|'(?:[^']|'')*'|"(?:[^"]|"")*"|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[\p{L}_$][\p{L}\p{N}_.$]*|<>|<=|>=|[!()+*/^,:=<>-]/guy;
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
    let token = tokens[pos++], sheetName;
    if (take("!")) {
      sheetName = token.startsWith("'") ? token.slice(1,-1).replaceAll("''","'") : token;
      token = tokens[pos++];
    }
    if (!token) fail("#VALUE!");
    if (token === "-" || token === "+") {
      const arg = primary(depth + 1);
      if (arg.op === "range") fail("#VALUE!");
      return { op: "sign", sign: token, arg };
    }
    if (token.startsWith('#')) return {op:'error',value:token};
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
    if (!take(":")) return { op: "ref", sheetName, address: ref.address };
    if (tokens[pos+1]==='!') {
      const endSheet=tokens[pos].replace(/^'|'$/g,'').replaceAll("''","'");
      if (endSheet.toUpperCase()!==sheetName?.toUpperCase()) fail("#NAME?");
      pos+=2;
    }
    const end = reference();
    const area = (Math.abs(ref.col - end.col) + 1) * (Math.abs(ref.row - end.row) + 1);
    if (area > MAX_RANGE) fail("#LIMIT!");
    return { op: "range", sheetName, start: ref, end };
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
  const states = book.sheets.map(() => new Map());
  const errors = book.sheets.map(() => new Set());
  const names = new Map(book.sheets.map((s,i)=>[s.name.toUpperCase(),i]));
  function calculate(sheetIndex, address) {
    const sheet = book.sheets[sheetIndex];
    if (!sheet) return "#REF!";
    if (states[sheetIndex].has(address)) return sheet.cells[address]?.saved ? sheet.cells[address].v : '';
    let current = sheetIndex, scenario = null;
    const resolve = name => name === undefined ? current : names.get(name.toUpperCase()) ?? -1;
    const unavailable = () => { throw new Unavailable(); };
    const visiting = new Set();
    let work = 0;
    function step() { if (++work > MAX_WORK) fail("#LIMIT!"); }
    function expressionValue(expression) {
      if (!parsed.has(expression)) {
        try { parsed.set(expression,parse(expression)); }
        catch (error) {
          if (book.xlsx && error instanceof FormulaError) unavailable();
          throw error;
        }
      }
      return evaluate(parsed.get(expression)) ?? 0;
    }
    function groupValue(group,key) {
      const [r,c]=coords(key);
      if (group.type === 'array') {
        if (group.source) {
          const {r:row,c:col,qualifier}=group.source;
          return expressionValue(qualifier+a1(row+r-group.top,col+c-group.left));
        }
        if (group.top===group.bottom && group.left===group.right) return expressionValue(group.expression);
        unavailable();
      }
      if (group.type!=='dataTable' || group.deleted || scenario || group.top<2 || group.left<2) unavailable();
      if (!coordinates(group.input1) || (group.two && !coordinates(group.input2))) unavailable();
      const header = a1(group.row || group.two ? r : group.top-1,
                        group.row || group.two ? group.left-1 : c);
      const formula = group.two ? a1(group.top-1,group.left-1) : header;
      if (!book.sheets[current].cells[formula]?.f) unavailable();
      const inputs = new Map();
      const topValue = () => read(a1(group.top-1,c));
      const leftValue = () => read(a1(r,group.left-1));
      inputs.set(group.input1,group.two || group.row ? topValue() : leftValue());
      if (group.two) inputs.set(group.input2,leftValue());
      scenario = {sheet:current,inputs};
      try { return read(formula); } finally { scenario=null; }
    }
    function read(key, index=current) {
      step();
      if (!book.sheets[index]) fail("#REF!");
      if (scenario?.sheet===index && scenario.inputs.has(key)) return scenario.inputs.get(key);
      const cache = scenario ? null : caches[index];
      if (cache?.has(key)) return cache.get(key);
      if (!scenario && states[index].has(key)) unavailable();
      const cell = book.sheets[index].cells[key];
      if (!cell) return null;
      if (!cell.group && typeof cell.f !== "string") {
        if (cell.error) fail(cell.v);
        return cell.v === "" ? null : cell.v ?? null;
      }
      const id=index+'!'+key+(scenario?'@':'');
      if (visiting.has(id)) fail("#CYCLE!");
      if (visiting.size >= MAX_DEPTH) fail("#LIMIT!");
      const previous=current;
      visiting.add(id); current=index;
      try {
        const value = cell.group ? groupValue(cell.group,key) : expressionValue(cell.f);
        cache?.set(key,value);
        return value;
      } catch(error) {
        if(cell.group && error instanceof FormulaError && ['#LIMIT!','#CYCLE!'].includes(error.message)) unavailable();
        throw error;
      } finally { visiting.delete(id); current=previous; }
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
            include(read(label + row,resolve(arg.sheetName)), true);
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
        case "error": fail(node.value);
        case "ref": return read(node.address,resolve(node.sheetName));
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
      const cell=sheet.cells[address];
      if (error instanceof Unavailable || (cell?.group && error instanceof FormulaError && ['#LIMIT!','#CYCLE!'].includes(error.message))) {
        states[sheetIndex].set(address,cell?.saved ? 'Last saved value' : 'No saved result');
        return cell?.saved ? cell.v : '';
      }
      if (error instanceof FormulaError) {errors[sheetIndex].add(address);return error.message;}
      if (error instanceof RangeError) return "#LIMIT!";
      throw error;
    }
  }
  calculate.isError = (index,key) => errors[index].has(key);
  calculate.status = (index,key) => {
    if(book.xlsx && !caches[index].has(key) && !states[index].has(key) && (book.sheets[index].cells[key]?.f || book.sheets[index].cells[key]?.group))calculate(index,key);
    return states[index].get(key) || '';
  };
  return calculate;
}
