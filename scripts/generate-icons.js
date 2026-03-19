const { createCanvas } = require("canvas");
const fs = require("fs");
const path = require("path");

const ASSETS_DIR = path.join(__dirname, "..", "assets");
const BG_COLOR = "#1a6b37";
const TEXT_COLOR = "#ffffff";

const icons = [
  { file: "icon-16.png", size: 16 },
  { file: "icon-32.png", size: 32 },
  { file: "icon-64.png", size: 64 },
  { file: "icon-80.png", size: 80 },
  { file: "icon-128.png", size: 128 },
  { file: "logo-filled.png", size: 128 },
];

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

for (const { file, size } of icons) {
  const canvas = createCanvas(size, size);
  const ctx = canvas.getContext("2d");

  // Rounded rectangle background
  const radius = Math.round(size * 0.18);
  roundRect(ctx, 0, 0, size, size, radius);
  ctx.fillStyle = BG_COLOR;
  ctx.fill();

  // "AI" text
  const fontSize = Math.round(size * 0.5);
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = `bold ${fontSize}px sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("AI", size / 2, size / 2 + size * 0.03);

  const out = path.join(ASSETS_DIR, file);
  fs.writeFileSync(out, canvas.toBuffer("image/png"));
  console.log(`  ${file} (${size}x${size})`);
}

console.log("Done.");
