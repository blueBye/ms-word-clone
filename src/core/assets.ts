export function imageData(dataUrl: string) {
  const match =
    /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/]+={0,2})$/.exec(dataUrl);
  if (!match || match[2].length > 28_000_000)
    throw new Error("Choose a PNG or JPEG image smaller than 20 MB.");
  const bytes = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0));
  if (
    match[1] === "image/png"
      ? ![137, 80, 78, 71, 13, 10, 26, 10].every((b, i) => bytes[i] === b)
      : bytes[0] !== 255 || bytes[1] !== 216 || bytes[2] !== 255
  )
    throw new Error("Image contents do not match the file type.");
  return {
    mime: match[1],
    extension: match[1] === "image/png" ? "png" : "jpg",
    bytes,
  };
}
export function hyperlink(href: string) {
  if (
    !/^(https?:\/\/|mailto:|#)/i.test(href) ||
    /[\u0000-\u0020]/.test(href) ||
    href.length > 4096
  )
    throw new Error(
      "Enter an https://, http://, mailto: or document bookmark address.",
    );
  return href;
}
