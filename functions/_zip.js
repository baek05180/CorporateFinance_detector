// GitHub Actions 결과물(zip)에서 파일 하나 꺼내기 — 무압축(0)·deflate(8) 지원
export async function unzipEntry(buf, name) {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('zip: no end record');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('zip: bad directory');
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const fname = new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen));
    if (fname === name) {
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      const data = u8.subarray(start, start + csize);
      if (method === 0) return data;
      if (method !== 8) throw new Error('zip: unsupported method');
      const ds = new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw')));
      return new Uint8Array(await ds.arrayBuffer());
    }
    p += 46 + nlen + xlen + clen;
  }
  return null;
}
