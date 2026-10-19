#!/usr/bin/env node
/**
 * Build the Microsoft 365 app package (zip) for the unified manifest.
 *
 *   ADDIN_PUBLIC_URL=https://addins.example.com/stepchart/ npm run build
 *   npm run package
 *
 * Produces dist/probability-step-chart.zip containing manifest.json and the two
 * app icons it references. Upload that zip in the Microsoft 365 admin center
 * (Integrated apps) or sideload it through Teams / Agents Toolkit. The XML
 * manifest (dist/manifest.xml) is distributed as-is for older Office versions.
 */
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const dist = join(root, "dist");
const manifestPath = join(dist, "manifest.json");
if (!existsSync(manifestPath)) {
  console.error("dist/manifest.json not found. Run `npm run build` first.");
  process.exit(1);
}

const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const urls = JSON.stringify(manifest).match(/https:\/\/[^"]+/g) ?? [];
if (urls.some((u) => u.startsWith("https://localhost"))) {
  console.warn(
    "Warning: the manifest still points at https://localhost:3000. Set ADDIN_PUBLIC_URL and rebuild before distributing.",
  );
}

const files = [
  ["manifest.json", readFileSync(manifestPath)],
  [manifest.icons.color, readFileSync(join(dist, manifest.icons.color))],
  [manifest.icons.outline, readFileSync(join(dist, manifest.icons.outline))],
];

// --- minimal ZIP writer (deflate), no dependencies -------------------------
const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

const local = [];
const central = [];
let offset = 0;
for (const [name, data] of files) {
  const nameBuf = Buffer.from(name, "utf8");
  const compressed = zlib.deflateRawSync(data);
  const crc = crc32(data);
  const header = Buffer.alloc(30);
  header.writeUInt32LE(0x04034b50, 0);
  header.writeUInt16LE(20, 4); // version needed
  header.writeUInt16LE(0x0800, 6); // UTF-8 names
  header.writeUInt16LE(8, 8); // deflate
  header.writeUInt32LE(0, 10); // time/date
  header.writeUInt32LE(crc, 14);
  header.writeUInt32LE(compressed.length, 18);
  header.writeUInt32LE(data.length, 22);
  header.writeUInt16LE(nameBuf.length, 26);
  header.writeUInt16LE(0, 28);
  local.push(header, nameBuf, compressed);

  const entry = Buffer.alloc(46);
  entry.writeUInt32LE(0x02014b50, 0);
  entry.writeUInt16LE(20, 4);
  entry.writeUInt16LE(20, 6);
  entry.writeUInt16LE(0x0800, 8);
  entry.writeUInt16LE(8, 10);
  entry.writeUInt32LE(0, 12);
  entry.writeUInt32LE(crc, 16);
  entry.writeUInt32LE(compressed.length, 20);
  entry.writeUInt32LE(data.length, 24);
  entry.writeUInt16LE(nameBuf.length, 28);
  entry.writeUInt32LE(offset, 42);
  central.push(entry, nameBuf);
  offset += header.length + nameBuf.length + compressed.length;
}
const centralBuf = Buffer.concat(central);
const end = Buffer.alloc(22);
end.writeUInt32LE(0x06054b50, 0);
end.writeUInt16LE(files.length, 8);
end.writeUInt16LE(files.length, 10);
end.writeUInt32LE(centralBuf.length, 12);
end.writeUInt32LE(offset, 16);

const out = join(dist, "probability-step-chart.zip");
writeFileSync(out, Buffer.concat([...local, centralBuf, end]));
console.log(`Wrote ${out} (${files.map(([n]) => n).join(", ")})`);
