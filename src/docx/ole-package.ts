import * as CFB from "cfb";
import { binaryData } from "../core/assets.js";
// Windows Object Packager stores arbitrary files in an OLE compound container.
const clsid = "0C00030000000000C000000000000046";
const join = (...parts: Uint8Array[]) => {
  const bytes = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let i = 0;
  for (const p of parts) {
    bytes.set(p, i);
    i += p.length;
  }
  return bytes;
};
const u32 = (n: number) => {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
};
const ansi = (s: string) =>
  Uint8Array.from(
    [...s.replace(/[^\x20-\x7e]/g, "_"), "\0"].map((c) => c.charCodeAt(0)),
  );
export function embeddedFile(
  asset: {
    dataUrl: string;
    name?: string;
    mime?: string;
    progId?: string;
    cells?: string[][];
  },
  bytes?: Uint8Array,
) {
  const data = bytes ?? binaryData(asset.dataUrl).bytes,
    name = (asset.name ?? "File.bin")
      .split(/[\\/]/)
      .at(-1)!
      .replace(/[\x00-\x1f]/g, "_");
  const extension = name.split(".").at(-1)!.toLowerCase();
  const office: Record<string, [string, string]> = {
    xlsx: [
      "Excel.Sheet.12",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ],
    docx: [
      "Word.Document.12",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ],
    pptx: [
      "PowerPoint.Show.12",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ],
  };
  if (office[extension] && data[0] === 80 && data[1] === 75)
    return {
      bytes: data,
      extension,
      mime: office[extension][1],
      kind: "package" as const,
      progId: office[extension][0],
    };
  if (
    data[0] === 0xd0 &&
    data[1] === 0xcf &&
    asset.progId &&
    asset.progId !== "Package"
  )
    return {
      bytes: data,
      extension: "bin",
      mime: "application/vnd.openxmlformats-officedocument.oleObject",
      kind: "oleObject" as const,
      progId: asset.progId,
    };
  const label = ansi(name),
    utf16 = new Uint8Array(name.length * 2);
  for (let i = 0; i < name.length; i++)
    new DataView(utf16.buffer).setUint16(i * 2, name.charCodeAt(i), true);
  const wide = join(u32(name.length), utf16),
    payload = join(
      new Uint8Array([2, 0]),
      label,
      label,
      new Uint8Array([0, 0, 3, 0]),
      u32(label.length),
      label,
      u32(data.length),
      data,
      wide,
      wide,
      wide,
    );
  const container = CFB.utils.cfb_new({ CLSID: clsid });
  CFB.utils.cfb_add(
    container,
    "\x01Ole10Native",
    join(u32(payload.length), payload),
  );
  const classBytes = Uint8Array.from(
    clsid.match(/../g)!.map((n) => parseInt(n, 16)),
  );
  CFB.utils.cfb_add(
    container,
    "\x01CompObj",
    join(
      u32(0xfffe0001),
      u32(0xa03),
      u32(0xffffffff),
      classBytes,
      u32(12),
      ansi("OLE Package"),
      u32(0),
      u32(8),
      ansi("Package"),
      u32(0x71b239f4),
      u32(0),
      u32(0),
      u32(0),
    ),
  );
  CFB.utils.cfb_add(
    container,
    "\x03ObjInfo",
    new Uint8Array([64, 0, 3, 0, 1, 0]),
  );
  return {
    bytes: new Uint8Array(
      CFB.write(container, { type: "array", fileType: "cfb" }),
    ),
    extension: "bin",
    mime: "application/vnd.openxmlformats-officedocument.oleObject",
    kind: "oleObject" as const,
    progId: "Package",
  };
}
export function unpackFile(
  bytes: Uint8Array,
): { name: string; bytes: Uint8Array } | undefined {
  if (bytes[0] !== 0xd0 || bytes[1] !== 0xcf) return;
  try {
    const entry = CFB.find(
      CFB.read(bytes, { type: "array" }),
      "\x01Ole10Native",
    );
    if (!entry) return;
    const raw = new Uint8Array(entry.content),
      view = new DataView(raw.buffer);
    let at = 6;
    const z = () => {
      const start = at;
      while (at < raw.length && raw[at]) at++;
      if (at === raw.length) throw new Error("Unterminated package name");
      return new TextDecoder("windows-1252").decode(raw.subarray(start, at++));
    };
    let name = z();
    z();
    at += 4;
    const length = view.getUint32(at, true);
    at += 4 + length;
    const size = view.getUint32(at, true);
    at += 4;
    if (size > raw.length - at) return;
    const data = raw.slice(at, at + size);
    at += size;
    if (at + 4 <= raw.length) {
      const chars = view.getUint32(at, true);
      if (chars < 32768 && at + 4 + chars * 2 <= raw.length)
        name = new TextDecoder("utf-16le").decode(
          raw.subarray(at + 4, at + 4 + chars * 2),
        );
    }
    return { name, bytes: data };
  } catch {
    return;
  }
}
