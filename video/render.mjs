// Video kompozisyonunu kare kare MP4'e çevirir: node render.mjs 01-226-meslek.html
// Her kare için sayfadaki window.render(t) çağrılır, ekran görüntüsü ffmpeg'e borulanır → cikti/<ad>.mp4
// Ayrıca 1. saniyeden kapak karesi: cikti/<ad>-kapak.jpg
import path from "node:path";
import { spawn } from "node:child_process";
import { mkdirSync } from "node:fs";
import { fileURLToPath, pathToFileURL } from "node:url";
import { chromium } from "playwright";
import ffmpeg from "ffmpeg-static";
import { synth } from "./synth.mjs";

const DIR = path.dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] ?? "01-226-meslek.html";
const ad = path.basename(file, ".html");
const OUT = path.join(DIR, "cikti");
mkdirSync(OUT, { recursive: true });

const b = await chromium.launch({ channel: "msedge", headless: true });
const pg = await b.newPage({ viewport: { width: 1080, height: 1920 } });
await pg.goto(pathToFileURL(path.join(DIR, file)).href + "?render", { waitUntil: "networkidle" });
await pg.evaluate(() => document.fonts.ready);
const { DUR, FPS, AUDIO } = await pg.evaluate(() => ({ DUR: window.DUR, FPS: window.FPS, AUDIO: window.AUDIO }));
const N = Math.round(DUR * FPS);

// ses: sayfa window.AUDIO tanımlıyorsa müzik + efektler üretilip videoya eklenir
const wav = path.join(OUT, ad + ".wav");
if (AUDIO) synth({ dur: DUR, ...AUDIO }, wav);
const sesArg = AUDIO ? ["-i", wav, "-c:a", "aac", "-b:a", "192k", "-shortest"] : [];

const ff = spawn(ffmpeg, ["-y", "-f", "image2pipe", "-framerate", String(FPS), "-c:v", "mjpeg", "-i", "-", ...sesArg,
  "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", "17", "-preset", "slow", "-movflags", "+faststart",
  path.join(OUT, ad + ".mp4")], { stdio: ["pipe", "ignore", "inherit"] });

for (let i = 0; i < N; i++) {
  await pg.evaluate(t => window.render(t), i / FPS);
  const buf = await pg.screenshot({ type: "jpeg", quality: 95 });
  if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once("drain", r));
  if (i % 60 === 0) process.stdout.write(`\r${i}/${N}`);
}
ff.stdin.end();
await new Promise(r => ff.on("close", r));

await pg.evaluate(() => window.render(2.0));
await pg.screenshot({ path: path.join(OUT, ad + "-kapak.jpg"), type: "jpeg", quality: 92 });
await b.close();
console.log(`\n✓ cikti/${ad}.mp4 (${N} kare, ${DUR} sn)`);
