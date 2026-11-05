export function imageData(dataUrl: string) {
  const match =
    /^data:(image\/(?:png|jpeg|gif|bmp|webp));base64,([A-Za-z0-9+/]+={0,2})$/.exec(
      dataUrl,
    );
  if (!match || match[2].length > 28_000_000)
    throw new Error("Choose a supported raster image smaller than 20 MB.");
  const bytes = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0));
  const signature = (offset: number, values: number[]) =>
    values.every((b, i) => bytes[offset + i] === b);
  if (
    !{
      "image/png": signature(0, [137, 80, 78, 71, 13, 10, 26, 10]),
      "image/jpeg": signature(0, [255, 216, 255]),
      "image/gif": signature(0, [71, 73, 70, 56]),
      "image/bmp": signature(0, [66, 77]),
      "image/webp":
        signature(0, [82, 73, 70, 70]) && signature(8, [87, 69, 66, 80]),
    }[match[1]]
  )
    throw new Error("Image contents do not match the file type.");
  return {
    mime: match[1],
    extension: match[1] === "image/jpeg" ? "jpg" : match[1].slice(6),
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
export function binaryData(dataUrl: string) {
  const match = /^data:([\w.+/-]+);base64,([A-Za-z0-9+/]+={0,2})$/.exec(
    dataUrl,
  );
  if (!match || match[2].length > 28_000_000)
    throw new Error("Choose a file smaller than 20 MB.");
  return {
    mime: match[1],
    bytes: Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0)),
  };
}
export function modelData(dataUrl: string) {
  const { bytes } = binaryData(dataUrl),
    view = new DataView(bytes.buffer);
  if (
    bytes.length < 24 ||
    view.getUint32(0, true) !== 0x46546c67 ||
    view.getUint32(4, true) !== 2 ||
    view.getUint32(8, true) !== bytes.length ||
    view.getUint32(16, true) !== 0x4e4f534a
  )
    throw new Error("Choose a GLB 2.0 model.");
  const length = view.getUint32(12, true);
  if (length > bytes.length - 20) throw new Error("Invalid GLB JSON chunk.");
  const json = JSON.parse(
    new TextDecoder().decode(bytes.subarray(20, 20 + length)),
  );
  for (const item of [...(json.buffers ?? []), ...(json.images ?? [])])
    if (item.uri && !/^data:/i.test(item.uri))
      throw new Error(
        "Use a self-contained GLB with embedded textures and buffers.",
      );
  return bytes;
}
export function videoUrl(value: string) {
  const url = new URL(value);
  if (url.protocol !== "https:")
    throw new Error("Use an HTTPS YouTube or Vimeo video URL.");
  if (
    [
      "www.youtube.com",
      "youtube.com",
      "youtu.be",
      "www.youtube-nocookie.com",
    ].includes(url.hostname)
  ) {
    const id =
      url.hostname === "youtu.be"
        ? url.pathname.slice(1)
        : url.pathname.startsWith("/embed/")
          ? url.pathname.split("/")[2]
          : url.searchParams.get("v");
    if (!id || !/^[\w-]{11}$/.test(id))
      throw new Error("Invalid YouTube video URL.");
    return `https://www.youtube-nocookie.com/embed/${id}`;
  }
  if (
    ["vimeo.com", "www.vimeo.com", "player.vimeo.com"].includes(url.hostname)
  ) {
    const id = url.pathname.split("/").filter(Boolean).at(-1);
    if (!id || !/^\d+$/.test(id)) throw new Error("Invalid Vimeo video URL.");
    return `https://player.vimeo.com/video/${id}`;
  }
  throw new Error("Use a YouTube or Vimeo video URL.");
}
