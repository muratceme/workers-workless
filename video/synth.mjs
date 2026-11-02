// Video sesini koddan üretir (telifsiz): arka plan müziği + sahneye senkron efektler → 48 kHz stereo WAV.
// Kullanım: synth({ dur, music: { drop, bpm, end }, cues: [[t, "hit"], ...] }, "cikti/x.wav")
import { writeFileSync } from "node:fs";

const SR = 48000;
const TAU = Math.PI * 2;

// deterministik gürültü (her render aynı sesi versin)
let seed = 1;
const noise = () => { seed = (seed * 16807) % 2147483647; return seed / 1073741823.5 - 1; };

// RBJ biquad
function biquad(type, f, q = .707) {
  const w = TAU * f / SR, c = Math.cos(w), s = Math.sin(w), a = s / (2 * q);
  let b0, b1, b2, a0 = 1 + a, a1 = -2 * c, a2 = 1 - a;
  if (type === "lp") { b0 = (1 - c) / 2; b1 = 1 - c; b2 = b0; }
  else if (type === "hp") { b0 = (1 + c) / 2; b1 = -(1 + c); b2 = b0; }
  else { b0 = a; b1 = 0; b2 = -a; } // bp
  let x1 = 0, x2 = 0, y1 = 0, y2 = 0;
  const f_ = { b0: b0 / a0, b1: b1 / a0, b2: b2 / a0, a1: a1 / a0, a2: a2 / a0 };
  return x => { const y = f_.b0 * x + f_.b1 * x1 + f_.b2 * x2 - f_.a1 * y1 - f_.a2 * y2; x2 = x1; x1 = x; y2 = y1; y1 = y; return y; };
}
// frekansı her örnekte değişen filtre için katsayıları yeniden hesaplayan sürüm
function sweepFilter(type, q = 1.2) {
  let x1 = 0, x2 = 0, y1 = 0, y2 = 0;
  return (x, f) => {
    const w = TAU * Math.min(f, SR * .45) / SR, c = Math.cos(w), s = Math.sin(w), a = s / (2 * q), a0 = 1 + a;
    let b0, b1, b2;
    if (type === "lp") { b0 = (1 - c) / 2; b1 = 1 - c; b2 = b0; } else { b0 = a; b1 = 0; b2 = -a; }
    const y = (b0 * x + b1 * x1 + b2 * x2 + 2 * c * y1 - (1 - a) * y2) / a0;
    x2 = x1; x1 = x; y2 = y1; y1 = y; return y;
  };
}

const midi = n => 440 * Math.pow(2, (n - 69) / 12);

export function synth({ dur, music, cues }, outPath) {
  const N = Math.ceil(dur * SR);
  const L = new Float32Array(N), R = new Float32Array(N);
  const add = (i, v, pan = 0) => { if (i >= 0 && i < N) { L[i] += v * (1 - pan) ; R[i] += v * (1 + pan); } };

  // ---------------- efektler ----------------
  const fx = {
    // derin darbe: düşen sinüs + gürültü patlaması
    hit(t, g = 1) {
      const n = SR * 1.2, lp = biquad("lp", 900); let ph = 0;
      for (let i = 0; i < n; i++) {
        const k = i / SR, f = 38 + 110 * Math.exp(-k * 18);
        ph += TAU * f / SR;
        const v = Math.sin(ph) * Math.exp(-k * 3.2) * .9 + lp(noise()) * Math.exp(-k * 14) * .8;
        add(Math.round(t * SR) + i, Math.tanh(v * 1.6) * g * .55);
      }
    },
    bigHit(t) { fx.hit(t, 1.25); fx.whoosh(t - .02, .5, .5, true); },
    // yukarı süpüren gerilim
    riser(t, d) {
      const n = Math.round(d * SR), f = sweepFilter("bp", 2.5); let ph = 0;
      for (let i = 0; i < n; i++) {
        const k = i / n, fr = 200 + 5000 * k * k;
        ph += TAU * (110 + 330 * k * k) / SR;
        const v = f(noise(), fr) * .9 + Math.sin(ph) * .12 * k;
        add(Math.round(t * SR) + i, v * k * k * .45, Math.sin(k * 20) * .2);
      }
    },
    // hava geçişi
    whoosh(t, d = .4, g = .35, down = false) {
      const n = Math.round(d * SR), f = sweepFilter("bp", 1.4);
      for (let i = 0; i < n; i++) {
        const k = i / n, env = Math.sin(Math.PI * Math.pow(k, .6));
        const fr = down ? 3000 - 2600 * k : 400 + 3600 * k;
        add(Math.round(t * SR) + i, f(noise(), fr) * env * g * 1.6, (k - .5) * .6);
      }
    },
    // dijital bozulma
    glitch(t, d) {
      const n = Math.round(d * SR); let hold = 0, v = 0, gate = 1;
      for (let i = 0; i < n; i++) {
        if (i % 2400 === 0) gate = noise() > -.2 ? 1 : 0;
        if (hold-- <= 0) { v = noise(); hold = 6 + (i % 7) * 3; }
        const s = (Math.sign(Math.sin(TAU * 180 * i / SR)) * .4 + v) * gate;
        add(Math.round(t * SR) + i, s * .16, noise() * .5);
      }
    },
    // çizgi çekme sesi
    swoosh(t) { fx.whoosh(t, .28, .5); },
    // kısa tık
    tick(t, g = .22, f = 2400) {
      const n = SR * .04;
      for (let i = 0; i < n; i++) { const k = i / SR; add(Math.round(t * SR) + i, Math.sin(TAU * f * k) * Math.exp(-k * 160) * g); }
    },
    // yumuşak "pop" (kart tiki, rozet)
    pop(t, g = .3, f = 880) {
      const n = SR * .18; let ph = 0;
      for (let i = 0; i < n; i++) {
        const k = i / SR; ph += TAU * (f * (1 + .6 * Math.exp(-k * 40))) / SR;
        add(Math.round(t * SR) + i, Math.sin(ph) * Math.exp(-k * 26) * g);
      }
    },
    // kapanış çanı
    chime(t) {
      const notes = [76, 81, 83, 88];
      notes.forEach((m, j) => {
        const n = SR * 2.5, f = midi(m), st = t + j * .07;
        for (let i = 0; i < n; i++) {
          const k = i / SR;
          const v = (Math.sin(TAU * f * k) + .3 * Math.sin(TAU * f * 2.01 * k) + .12 * Math.sin(TAU * f * 3.98 * k)) * Math.exp(-k * 2.2) * Math.min(1, k * 300);
          add(Math.round(st * SR) + i, v * .09, (j - 1.5) * .25);
        }
      });
    },
  };
  for (const [t, name, ...a] of cues) fx[name](t, ...a);

  // ---------------- müzik ----------------
  if (music) {
    const { drop, bpm = 120, end = dur } = music;
    const beat = 60 / bpm;
    // akorlar (La minör: Am – F – C – G), her biri 4 vuruş
    const CH = [[57, 60, 64], [53, 57, 60], [48, 52, 55], [55, 59, 62]];
    const ROOT = [45, 41, 48, 43];
    const vol = t => t < drop ? 0 : Math.min(1, (end - t) / 1.6);
    const pre = t => t < drop ? Math.min(1, t / 1.5) * .7 : 0;

    // drop öncesi: düşük gerilim dronu (La)
    let dph = 0;
    for (let i = 0; i < Math.round(drop * SR); i++) {
      const t = i / SR; dph += TAU * midi(33) / SR;
      const v = (Math.sin(dph) + .4 * Math.sin(dph * 2.003) + .2 * Math.sin(dph * 3)) * pre(t) * .12 * (1 + .25 * Math.sin(TAU * 4 * t));
      add(i, v);
    }

    const beats = Math.floor((end - drop) / beat);
    for (let b = 0; b < beats; b++) {
      const t0 = drop + b * beat, ci = Math.floor(b / 4) % 4, g = vol(t0);
      if (g <= 0) continue;
      // kick
      { let ph = 0; for (let i = 0; i < SR * .35; i++) { const k = i / SR; ph += TAU * (45 + 120 * Math.exp(-k * 30)) / SR; add(Math.round(t0 * SR) + i, Math.sin(ph) * Math.exp(-k * 9) * .5 * g); } }
      // clap (2 ve 4)
      if (b % 2 === 1) { const bp = biquad("bp", 1500, .9); for (let i = 0; i < SR * .18; i++) { const k = i / SR; add(Math.round(t0 * SR) + i, bp(noise()) * Math.exp(-k * 22) * .55 * g, .1); } }
      // hi-hat (ara vuruş + 16'lıklar hafif)
      for (const [off, hg] of [[.5, .2], [.25, .06], [.75, .06]]) {
        const hp = biquad("hp", 8000); const st = Math.round((t0 + off * beat) * SR);
        for (let i = 0; i < SR * .05; i++) add(st + i, hp(noise()) * Math.exp(-i / SR * 70) * hg * g, -.3);
      }
      // bas: 8'lik, yan zincir hissi için vuruşun başında kısık
      for (const off of [0, .5]) {
        const st = Math.round((t0 + off * beat) * SR), f = midi(ROOT[ci] - 12), lp = biquad("lp", 380);
        let ph = 0;
        for (let i = 0; i < beat * .5 * SR; i++) {
          const k = i / SR; ph += TAU * f / SR;
          const duck = off === 0 ? Math.min(1, k * 12) : 1;
          add(st + i, lp(Math.sin(ph) + .5 * Math.sign(Math.sin(ph))) * Math.exp(-k * 3) * duck * .26 * g);
        }
      }
      // pad: her akorun başında 4 vuruşluk, detune testere + lowpass
      if (b % 4 === 0) {
        const st = Math.round(t0 * SR), n = Math.round(beat * 4 * SR);
        CH[ci].forEach((m, j) => {
          const lp = biquad("lp", 1400); const fs = [midi(m) * .997, midi(m) * 1.003];
          for (let i = 0; i < n; i++) {
            const k = i / SR, env = Math.min(1, k * 4) * Math.min(1, (n - i) / (SR * .3));
            let s = 0; for (const f of fs) s += 2 * ((f * k) % 1) - 1;
            const v = lp(s) * env * .045 * g * (1 - .45 * Math.exp(-((k % beat)) * 8)); // vuruşa göre nefes
            add(st + i, v, (j - 1) * .35);
          }
        });
      }
      // arp: 16'lık parıltı, yalnız ikinci yarıda
      if (t0 > drop + beat * 8) {
        const notes = CH[ci].map(m => m + 12);
        for (let s16 = 0; s16 < 4; s16++) {
          const st = Math.round((t0 + s16 * beat / 4) * SR), f = midi(notes[(b * 4 + s16) % 3]);
          for (let i = 0; i < SR * .12; i++) { const k = i / SR; add(st + i, Math.sin(TAU * f * k) * Math.exp(-k * 30) * .05 * g, s16 % 2 ? .4 : -.4); }
        }
      }
    }
  }

  // ---------------- miks: yumuşak sınırlama + normalize ----------------
  let peak = 0;
  for (let i = 0; i < N; i++) { L[i] = Math.tanh(L[i] * 1.1); R[i] = Math.tanh(R[i] * 1.1); peak = Math.max(peak, Math.abs(L[i]), Math.abs(R[i])); }
  const gain = .89 / (peak || 1);
  const buf = Buffer.alloc(44 + N * 4);
  buf.write("RIFF", 0); buf.writeUInt32LE(36 + N * 4, 4); buf.write("WAVEfmt ", 8);
  buf.writeUInt32LE(16, 16); buf.writeUInt16LE(1, 20); buf.writeUInt16LE(2, 22); buf.writeUInt32LE(SR, 24);
  buf.writeUInt32LE(SR * 4, 28); buf.writeUInt16LE(4, 32); buf.writeUInt16LE(16, 34); buf.write("data", 36); buf.writeUInt32LE(N * 4, 40);
  for (let i = 0; i < N; i++) {
    buf.writeInt16LE(Math.round(Math.max(-1, Math.min(1, L[i] * gain)) * 32767), 44 + i * 4);
    buf.writeInt16LE(Math.round(Math.max(-1, Math.min(1, R[i] * gain)) * 32767), 46 + i * 4);
  }
  writeFileSync(outPath, buf);
}
