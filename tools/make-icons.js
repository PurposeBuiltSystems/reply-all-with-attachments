/*
 * Generate placeholder ribbon icons (16/32/80 px PNGs) with no third-party
 * deps — just Node's zlib + Buffer. Draws a rounded blue tile with a white
 * "reply" arrow so the button is recognizable. Swap these for real artwork
 * before publishing.
 *
 *   node tools/make-icons.js
 */

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const BG = [15, 108, 189]; // Office blue
const FG = [255, 255, 255]; // arrow

// --- CRC32 (PNG chunk checksum) -------------------------------------------
const CRC_TABLE = (function () {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuf = Buffer.from(type, "ascii");
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

// --- Pixel drawing ---------------------------------------------------------
function makePng(size) {
  const px = Buffer.alloc(size * size * 4); // RGBA
  const r = Math.max(2, Math.round(size * 0.18)); // corner radius
  const c = size / 2;

  function set(x, y, rgb, a) {
    const i = (y * size + x) * 4;
    px[i] = rgb[0];
    px[i + 1] = rgb[1];
    px[i + 2] = rgb[2];
    px[i + 3] = a;
  }

  // Rounded-rect background.
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const inCorner =
        (x < r && y < r && (x - r) ** 2 + (y - r) ** 2 > r * r) ||
        (x >= size - r && y < r && (x - (size - r - 1)) ** 2 + (y - r) ** 2 > r * r) ||
        (x < r && y >= size - r && (x - r) ** 2 + (y - (size - r - 1)) ** 2 > r * r) ||
        (x >= size - r && y >= size - r && (x - (size - r - 1)) ** 2 + (y - (size - r - 1)) ** 2 > r * r);
      set(x, y, BG, inCorner ? 0 : 255);
    }
  }

  // A chunky white left-pointing reply arrow.
  const th = Math.max(1, Math.round(size * 0.09)); // stroke thickness
  const ax0 = Math.round(size * 0.26); // arrow tip x
  const ax1 = Math.round(size * 0.66); // shaft right x
  const head = Math.round(size * 0.16); // arrowhead half-height
  for (let x = ax0; x <= ax1; x++) {
    for (let t = -th; t <= th; t++) set(x, Math.round(c) + t, FG, 255); // shaft
  }
  for (let d = 0; d <= head; d++) {
    for (let t = -th; t <= th; t++) {
      set(ax0 + d, Math.round(c) - d + t, FG, 255); // upper diagonal
      set(ax0 + d, Math.round(c) + d + t, FG, 255); // lower diagonal
    }
  }
  // Short vertical riser at the shaft end (the "reply" hook).
  for (let y = Math.round(c); y <= Math.round(size * 0.72); y++) {
    for (let t = -th; t <= th; t++) set(ax1 + t, y, FG, 255);
  }

  // Pack scanlines with filter byte 0, then deflate.
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    px.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const outDir = path.join(__dirname, "..", "assets");
fs.mkdirSync(outDir, { recursive: true });
[16, 32, 80].forEach(function (size) {
  const file = path.join(outDir, "icon-" + size + ".png");
  fs.writeFileSync(file, makePng(size));
  console.log("wrote " + file);
});
