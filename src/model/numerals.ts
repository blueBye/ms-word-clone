/** Display-only digit shaping. Stored source text is never rewritten. */
export function localizeDigits(value: string, locale?: string): string {
  if (!locale || !/^(fa|ar|ur)(-|$)/i.test(locale)) return value;
  const digits = /^fa|^ur/i.test(locale) ? '۰۱۲۳۴۵۶۷۸۹' : '٠١٢٣٤٥٦٧٨٩';
  return value.replace(/[0-9]/g, (digit) => digits[Number(digit)]);
}
