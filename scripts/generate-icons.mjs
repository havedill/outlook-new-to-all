import { writeFileSync, mkdirSync } from "fs";
import { deflateSync } from "zlib";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const outDir = join(__dirname, "..", "docs", "assets");
mkdirSync(outDir, { recursive: true });

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuf = Buffer.from(type);
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const crcBuf = Buffer.alloc(4);
  const crcData = Buffer.concat([typeBuf, data]);
  crcBuf.writeUInt32BE(crc32(crcData));
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function insideEnvelope(x, y, size) {
  const cx = size / 2;
  const top = size * 0.28;
  const bottom = size * 0.72;
  const left = size * 0.22;
  const right = size * 0.78;
  if (y < top || y > bottom) return false;
  if (x < left || x > right) return false;
  const midY = (top + bottom) / 2;
  const slope = (bottom - top) / (right - left);
  const distFromFold =
    Math.abs(y - top - slope * Math.abs(x - cx)) < size * 0.06 && y < midY;
  const inBody = y >= midY - size * 0.02;
  const inFlap =
    y < midY &&
    y >= top + slope * Math.abs(x - cx) - size * 0.04 &&
    y <= top + slope * (right - left) / 2 + size * 0.02;
  return inBody || inFlap || distFromFold;
}

function makePng(size) {
  const raw = Buffer.alloc((size * size * 4) + size);
  let offset = 0;
  for (let y = 0; y < size; y++) {
    raw[offset++] = 0;
    for (let x = 0; x < size; x++) {
      const inIcon = insideEnvelope(x, y, size);
      const bgR = 0x00;
      const bgG = 0x78;
      const bgB = 0xd4;
      const fgR = 255;
      const fgG = 255;
      const fgB = 255;
      raw[offset++] = inIcon ? fgR : bgR;
      raw[offset++] = inIcon ? fgG : bgG;
      raw[offset++] = inIcon ? fgB : bgB;
      raw[offset++] = 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

for (const size of [16, 32, 64, 80, 128]) {
  writeFileSync(join(outDir, `icon-${size}.png`), makePng(size));
}

console.log("Wrote icons to", outDir);
