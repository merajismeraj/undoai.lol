/**
 * A minimal ZIP reader/writer, because a .docx and a .pptx are both just a
 * ZIP of XML parts.
 *
 * Hand-rolled rather than pulled from npm for one reason: this whole tool runs
 * in the visitor's browser, so a document never leaves their machine. Shipping
 * a general-purpose zip library to every page load to read a handful of
 * headers is a worse trade than 150 lines we control. Compression itself is
 * the platform's (`CompressionStream`), not ours.
 *
 * Deliberately not supported: ZIP64, encrypted entries, multi-disk archives.
 * Office never produces them for documents of the size this tool accepts, and
 * silently mangling one would be worse than refusing it.
 */

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;

/** A fixed DOS timestamp (1 Jan 2000) stamped on every entry we write. Real
    mtimes are a fingerprint of their own — when the file was generated, how
    fast the parts were written — so every rebuilt archive gets the same one. */
const DOS_DATE = ((2000 - 1980) << 9) | (1 << 5) | 1;
const DOS_TIME = 0;

export type ZipEntry = {
  name: string;
  /** 0 = stored, 8 = deflate. */
  method: number;
  crc32: number;
  /** Raw bytes exactly as they sit in the archive, still compressed. Kept so
      untouched parts (images, fonts, media) can be copied through verbatim
      instead of being decompressed and re-deflated for nothing. */
  compressed: Uint8Array;
  uncompressedSize: number;
};

export class ZipError extends Error {}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

export function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** `CompressionStream`/`DecompressionStream` are typed as accepting
    `BufferSource`, which `pipeThrough` won't unify with a `Uint8Array` source,
    hence the single cast here rather than at both call sites. */
async function runStream(bytes: Uint8Array, stream: GenericTransformStream) {
  const source = new Blob([bytes as BlobPart]).stream() as ReadableStream<Uint8Array>;
  const piped = source.pipeThrough(
    stream as unknown as ReadableWritablePair<Uint8Array, Uint8Array>,
  );
  const out = await new Response(piped).arrayBuffer();
  return new Uint8Array(out);
}

export function inflateRaw(bytes: Uint8Array): Promise<Uint8Array> {
  return runStream(bytes, new DecompressionStream("deflate-raw"));
}

export function deflateRaw(bytes: Uint8Array): Promise<Uint8Array> {
  return runStream(bytes, new CompressionStream("deflate-raw"));
}

/** Whether this browser can do raw deflate both ways. Checked before a file is
    ever read, so an unsupported browser gets a sentence rather than a stack. */
export function supportsZip(): boolean {
  return typeof DecompressionStream === "function" && typeof CompressionStream === "function";
}

/** Reads the central directory — the archive's own index — rather than walking
    local headers, since only the central directory is authoritative when an
    entry was written with a streaming data descriptor. */
export function readZip(buffer: ArrayBuffer): ZipEntry[] {
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer);

  let eocd = -1;
  const scanFrom = Math.max(0, bytes.length - 66_000);
  for (let i = bytes.length - 22; i >= scanFrom; i--) {
    if (view.getUint32(i, true) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new ZipError("Not a valid .docx or .pptx file.");

  const count = view.getUint16(eocd + 10, true);
  const cdOffset = view.getUint32(eocd + 16, true);
  if (count === 0xffff || cdOffset === 0xffffffff) {
    throw new ZipError("ZIP64 archives aren't supported. Re-save the file from Word or PowerPoint.");
  }

  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  let p = cdOffset;

  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== SIG_CENTRAL) throw new ZipError("Corrupt archive index.");

    const flags = view.getUint16(p + 8, true);
    if (flags & 0x1) throw new ZipError("The file is password-protected. Remove the password first.");

    const method = view.getUint16(p + 10, true);
    const crc = view.getUint32(p + 16, true);
    const compSize = view.getUint32(p + 20, true);
    const uncompSize = view.getUint32(p + 24, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));

    if (view.getUint32(localOffset, true) !== SIG_LOCAL) throw new ZipError("Corrupt archive entry.");
    const localNameLen = view.getUint16(localOffset + 26, true);
    const localExtraLen = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + localNameLen + localExtraLen;

    entries.push({
      name,
      method,
      crc32: crc,
      uncompressedSize: uncompSize,
      compressed: bytes.slice(dataStart, dataStart + compSize),
    });

    p += 46 + nameLen + extraLen + commentLen;
  }

  return entries;
}

export async function readEntry(entry: ZipEntry): Promise<Uint8Array> {
  if (entry.method === 0) return entry.compressed;
  if (entry.method === 8) return inflateRaw(entry.compressed);
  throw new ZipError(`Unsupported compression in "${entry.name}".`);
}

/** Re-deflates a part we rewrote, falling back to stored if deflate somehow
    made it bigger (tiny XML parts occasionally do). */
export async function makeEntry(name: string, data: Uint8Array): Promise<ZipEntry> {
  const deflated = await deflateRaw(data);
  const useDeflate = deflated.length < data.length;
  return {
    name,
    method: useDeflate ? 8 : 0,
    crc32: crc32(data),
    uncompressedSize: data.length,
    compressed: useDeflate ? deflated : data,
  };
}

export function writeZip(entries: ZipEntry[]): Blob {
  const encoder = new TextEncoder();
  const parts: BlobPart[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.name);

    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, SIG_LOCAL, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0, true); // flags — no data descriptor, sizes are known
    lv.setUint16(8, entry.method, true);
    lv.setUint16(10, DOS_TIME, true);
    lv.setUint16(12, DOS_DATE, true);
    lv.setUint32(14, entry.crc32, true);
    lv.setUint32(18, entry.compressed.length, true);
    lv.setUint32(22, entry.uncompressedSize, true);
    lv.setUint16(26, name.length, true);
    lv.setUint16(28, 0, true);
    local.set(name, 30);

    parts.push(local as BlobPart, entry.compressed as BlobPart);

    const dir = new Uint8Array(46 + name.length);
    const dv = new DataView(dir.buffer);
    dv.setUint32(0, SIG_CENTRAL, true);
    dv.setUint16(4, 20, true); // version made by
    dv.setUint16(6, 20, true); // version needed
    dv.setUint16(8, 0, true);
    dv.setUint16(10, entry.method, true);
    dv.setUint16(12, DOS_TIME, true);
    dv.setUint16(14, DOS_DATE, true);
    dv.setUint32(16, entry.crc32, true);
    dv.setUint32(20, entry.compressed.length, true);
    dv.setUint32(24, entry.uncompressedSize, true);
    dv.setUint16(28, name.length, true);
    dv.setUint32(42, offset, true);
    dir.set(name, 46);
    central.push(dir);

    offset += local.length + entry.compressed.length;
  }

  const centralSize = central.reduce((sum, d) => sum + d.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, SIG_EOCD, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  return new Blob([...parts, ...(central as BlobPart[]), eocd as BlobPart], {
    type: "application/zip",
  });
}
