import { FunctionPlugin, FunctionArgumentType, CellError, ErrorType } from "hyperformula";

/** Flatten a HyperFormula range argument into a 1D list of raw values. */
function flatValues(range) {
  if (!range) return [];
  if (Array.isArray(range)) return range.flat();
  if (range.data) return range.data.flat();
  return [range];
}

/** Excel-style criteria matching: numbers, text, comparison operators, wildcards. */
export function matchesCriterion(value, criterion) {
  if (criterion === null || criterion === undefined) return value === "" || value === null;
  const raw = typeof criterion === "string" ? criterion.trim() : criterion;
  if (typeof raw === "string" && raw.startsWith('"') && raw.endsWith('"')) return String(value ?? "") === raw.slice(1, -1);
  if (typeof raw !== "string") return Number(value) === Number(raw);
  const opMatch = /^(<=|>=|<>|=|<|>)(.*)$/.exec(raw);
  const op = opMatch ? opMatch[1] : "=";
  const operand = opMatch ? opMatch[2].trim() : raw;
  const operandNumber = Number(operand);
  const valueNumber = typeof value === "number" ? value : Number(value);
  const numeric = operand !== "" && Number.isFinite(operandNumber);
  if (numeric && Number.isFinite(valueNumber)) {
    switch (op) {
      case ">": return valueNumber > operandNumber;
      case "<": return valueNumber < operandNumber;
      case ">=": return valueNumber >= operandNumber;
      case "<=": return valueNumber <= operandNumber;
      case "<>": return valueNumber !== operandNumber;
      default: return valueNumber === operandNumber;
    }
  }
  const text = String(value ?? "").toLowerCase();
  const needle = operand.toLowerCase().replace(/^"|"$/g, "");
  if (op === "<>") return text !== needle;
  if (needle.includes("*") || needle.includes("?")) {
    const pattern = new RegExp("^" + needle.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$");
    return pattern.test(text);
  }
  if (op === "=" || op === "") return text === needle;
  return text === needle;
}

function numericValues(values) {
  return values
    .filter((value) => value !== null && value !== undefined && value !== "" && typeof value !== "boolean")
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
}

/** Extra Excel functions HyperFormula does not ship with. */
export class ExcelExtrasPlugin extends FunctionPlugin {
  averageifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("AVERAGEIFS"), (sumRange, ...rest) => {
      const sums = flatValues(sumRange);
      const conditions = [];
      for (let index = 0; index + 1 < rest.length; index += 2) {
        conditions.push({ values: flatValues(rest[index]), criterion: rest[index + 1] });
      }
      const selected = [];
      for (let index = 0; index < sums.length; index++) {
        if (conditions.every((condition) => matchesCriterion(condition.values[index], condition.criterion))) {
          selected.push(sums[index]);
        }
      }
      const numbers = numericValues(selected);
      if (!numbers.length) return new CellError(ErrorType.DIV_BY_ZERO);
      return numbers.reduce((total, value) => total + value, 0) / numbers.length;
    });
  }

  rank(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("RANK"), (value, range, order) => {
      const numbers = numericValues(flatValues(range));
      const target = Number(value);
      if (!Number.isFinite(target) || !numbers.length) return new CellError(ErrorType.VALUE);
      const ascending = Number(order) !== 0;
      const sorted = [...numbers].sort((left, right) => (ascending ? left - right : right - left));
      const position = sorted.findIndex((candidate) => candidate === target);
      return position < 0 ? new CellError(ErrorType.VALUE) : position + 1;
    });
  }

  mode(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("MODE"), (...ranges) => {
      const numbers = numericValues(ranges.flatMap((range) => flatValues(range)));
      const counts = new Map();
      for (const number of numbers) counts.set(number, (counts.get(number) ?? 0) + 1);
      let best = null;
      let bestCount = 1;
      for (const [number, count] of counts) {
        if (count > bestCount) { best = number; bestCount = count; }
      }
      return best === null ? new CellError(ErrorType.NA) : best;
    });
  }

  typeOf(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TYPE"), (value) => {
      if (value instanceof CellError) return 16;
      if (Array.isArray(value) || (value && typeof value === "object" && value.data)) return 64;
      if (typeof value === "number") return 1;
      if (typeof value === "string") return 2;
      if (typeof value === "boolean") return 4;
      return 1;
    });
  }

  concat(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("CONCAT"), (...args) =>
      args.flatMap((arg) => flatValues(arg)).map((value) => (value === null || value === undefined ? "" : String(value))).join(""));
  }

  textjoin(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TEXTJOIN"), (delimiter, ignoreEmpty, ...args) => {
      const parts = args
        .flatMap((arg) => flatValues(arg))
        .map((value) => (value === null || value === undefined ? "" : String(value)))
        .filter((value) => !ignoreEmpty || value !== "");
      return parts.join(delimiter === null || delimiter === undefined ? "" : String(delimiter));
    });
  }

  ifs(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("IFS"), (...args) => {
      for (let index = 0; index + 1 < args.length; index += 2) {
        const condition = args[index];
        if (condition instanceof CellError) return condition;
        if (condition) return args[index + 1];
      }
      return new CellError(ErrorType.NA);
    });
  }
}

ExcelExtrasPlugin.implementedFunctions = {
  AVERAGEIFS: {
    method: "averageifs",
    parameters: [
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.STRING },
    ],
    repeatLastArgs: 2,
  },
  RANK: {
    method: "rank",
    parameters: [
      { argumentType: FunctionArgumentType.NUMBER },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
    ],
  },
  MODE: {
    method: "mode",
    parameters: [{ argumentType: FunctionArgumentType.RANGE }],
    repeatLastArgs: 0,
  },
  TYPE: {
    method: "typeOf",
    parameters: [{ argumentType: FunctionArgumentType.SCALAR }],
  },
  CONCAT: {
    method: "concat",
    parameters: [{ argumentType: FunctionArgumentType.SCALAR }],
    repeatLastArgs: 1,
  },
  TEXTJOIN: {
    method: "textjoin",
    parameters: [
      { argumentType: FunctionArgumentType.STRING },
      { argumentType: FunctionArgumentType.BOOLEAN },
      { argumentType: FunctionArgumentType.SCALAR },
    ],
    repeatLastArgs: 1,
  },
  IFS: {
    method: "ifs",
    parameters: [
      { argumentType: FunctionArgumentType.BOOLEAN },
      { argumentType: FunctionArgumentType.SCALAR },
    ],
    repeatLastArgs: 2,
  },
};

ExcelExtrasPlugin.aliases = { "RANK.EQ": "RANK" };

export const excelExtrasTranslations = {
  enGB: {
    AVERAGEIFS: "AVERAGEIFS",
    RANK: "RANK",
    "RANK.EQ": "RANK.EQ",
    MODE: "MODE",
    TYPE: "TYPE",
    CONCAT: "CONCAT",
    TEXTJOIN: "TEXTJOIN",
    IFS: "IFS",
  },
};
