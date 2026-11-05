export type MathNode =
  | { kind: "text"; text: string }
  | { kind: "row"; items: MathNode[] }
  | { kind: "frac" | "sup" | "sub"; a: MathNode; b: MathNode }
  | { kind: "sqrt"; a: MathNode };
const symbols: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  theta: "θ",
  lambda: "λ",
  mu: "μ",
  pi: "π",
  sigma: "σ",
  omega: "ω",
  Omega: "Ω",
  Delta: "Δ",
  infty: "∞",
  times: "×",
  cdot: "·",
  pm: "±",
  le: "≤",
  ge: "≥",
  ne: "≠",
  approx: "≈",
  sum: "∑",
  int: "∫",
  rightarrow: "→",
  leftarrow: "←",
  partial: "∂",
  nabla: "∇",
  forall: "∀",
  exists: "∃",
  in: "∈",
};
export function parseEquation(input: string): MathNode {
  if (input.length > 10000) throw new Error("Equation is too long.");
  let i = 0,
    depth = 0;
  const atom = (): MathNode => {
    if (++depth > 50) throw new Error("Equation is nested too deeply.");
    let result: MathNode;
    if (input[i] === "{") {
      i++;
      result = row(true);
    } else if (input[i] === "\\") {
      i++;
      const name = /^[A-Za-z]+/.exec(input.slice(i))?.[0] ?? input[i++] ?? "";
      if (/^[A-Za-z]+$/.test(name)) i += name.length;
      if (name === "frac") result = { kind: "frac", a: atom(), b: atom() };
      else if (name === "sqrt") result = { kind: "sqrt", a: atom() };
      else if (symbols[name]) result = { kind: "text", text: symbols[name] };
      else if (name === "left" || name === "right") result = atom();
      else if (name === " " || name === ",")
        result = { kind: "text", text: " " };
      else throw new Error(`Unsupported equation command: \\${name}`);
    } else result = { kind: "text", text: input[i++] ?? "" };
    depth--;
    return result;
  };
  const row = (group = false): MathNode => {
    const items: MathNode[] = [];
    while (i < input.length && input[i] !== "}") {
      let value = atom();
      while (input[i] === "^" || input[i] === "_") {
        const kind = input[i++] === "^" ? "sup" : "sub";
        value = { kind, a: value, b: atom() };
      }
      items.push(value);
    }
    if (group) {
      if (input[i] !== "}") throw new Error("Close the equation's brace.");
      i++;
    }
    return { kind: "row", items };
  };
  const result = row();
  if (i < input.length) throw new Error("Unexpected closing brace.");
  return result;
}
