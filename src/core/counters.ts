export function formatCounter(value: number, format: string): string {
  if (format === "decimalZero") return String(value).padStart(2, "0");
  if (format === "lowerLetter" || format === "upperLetter") {
    let text = "";
    for (let n = value; n > 0; n = Math.floor((n - 1) / 26))
      text = String.fromCharCode(65 + ((n - 1) % 26)) + text;
    return format === "lowerLetter" ? text.toLowerCase() : text;
  }
  if (format === "upperRoman" || format === "lowerRoman") {
    let text = "",
      n = value;
    for (const [amount, symbol] of [
      [1000, "M"],
      [900, "CM"],
      [500, "D"],
      [400, "CD"],
      [100, "C"],
      [90, "XC"],
      [50, "L"],
      [40, "XL"],
      [10, "X"],
      [9, "IX"],
      [5, "V"],
      [4, "IV"],
      [1, "I"],
    ] as const)
      while (n >= amount) {
        text += symbol;
        n -= amount;
      }
    return format === "lowerRoman" ? text.toLowerCase() : text;
  }
  return String(value);
}
