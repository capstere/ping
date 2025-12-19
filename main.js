(() => {
  "use strict";

  // ---------- Safe helpers ----------
  const $ = (id) => document.getElementById(id);
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const lerp = (a, b, t) => a + (b - a) * t;

  function nowMs(){ return performance.now ? performance.now() : Date.now(); }

  // ---------- Countdown (7 Jan 2026 06:00 CET = 05:00 UTC) ----------
  const countdownEl = $("countdownValue");
  const TARGET_UTC_MS = Date.UTC(2026, 0, 7, 5, 0, 0);
  function pad2(n){ return String(n).padStart(2, "0"); }
  function tickCountdown(){
    const diff = TARGET_UTC_MS - Date.now();
    if (!countdownEl) return;
    if (diff <= 0){ countdownEl.textContent = "NU. ☕"; return; }
    const total = Math.floor(diff / 1000);
    const d = Math.floor(total / 86400);
    const h = Math.floor((total % 86400) / 3600);
    const m = Math.floor((total % 3600) / 60);
    const s = total % 60;
    countdownEl.textContent = `${d}d ${pad2(h)}:${pad2(m)}:${pad2(s)}`;
  }
  setInterval(tickCountdown, 1000);
  tickCountdown();

  // ---------- UI ----------
  const canvas = $("game");
  const toastEl = $("toast");
  const bubbleEl = $("bubble");
  const soundBtn = $("soundBtn");
  const helpBtn = $("helpBtn");
  const helpModal = $("helpModal");
  const closeHelpBtn = $("closeHelpBtn");
  const kickBtn = $("kickBtn");
  const sitBtn = $("sitBtn");
  const wonder = $("wonder");
  const closeWonderBtn = $("closeWonderBtn");
  const wonderImg = $("wonderImg");
  const tavlaImg = $("tavlaImg");
  const wonderFallback = $("wonderFallback");

  const moodEl = $("mood");
  const hintBtn = $("hintBtn");
  const achBtn  = $("achBtn");
  const achModal = $("achModal");
  const closeAchBtn = $("closeAchBtn");
  const boopCountEl = $("boopCount");
  const runTimeEl = $("runTime");
  const bestTimeEl = $("bestTime");
  const doorLink = $("doorLink");
  const highFiveBtn = $("highFiveBtn");
  const danceBtn = $("danceBtn");
  const confettiCanvas = $("confetti");

  // ---------- Mobile viewport helpers ----------
  function syncViewportVars(){
    // vh unit that follows the visible viewport (helps on iOS Safari)
    const vv = window.visualViewport;
    const h = vv ? vv.height : window.innerHeight;
    document.documentElement.style.setProperty("--vh", (h * 0.01) + "px");

    // Keep --topbar-h synced to the real header height (wraps on small screens)
    const topbar = document.querySelector(".topbar");
    if (topbar){
      const tbH = Math.ceil(topbar.getBoundingClientRect().height);
      document.documentElement.style.setProperty("--topbar-h", tbH + "px");
    }
  }

  window.addEventListener("resize", syncViewportVars, { passive:true });
  window.addEventListener("orientationchange", syncViewportVars, { passive:true });
  if (window.visualViewport){
    window.visualViewport.addEventListener("resize", syncViewportVars, { passive:true });
    window.visualViewport.addEventListener("scroll", syncViewportVars, { passive:true });
  }
  // Run once right away (and again shortly after layout/fonts settle)
  syncViewportVars();
  setTimeout(syncViewportVars, 80);




  function toast(msg, ms=1200){
    if (!toastEl) return;
    toastEl.textContent = msg;
    toastEl.classList.remove("hidden");
    clearTimeout(toast._t);
    toast._t = setTimeout(()=>toastEl.classList.add("hidden"), ms);
  }
  function bubble(msg, ms=1600){
    if (!bubbleEl) return;
    bubbleEl.textContent = msg;
    bubbleEl.classList.remove("hidden");
    clearTimeout(bubble._t);
    bubble._t = setTimeout(()=>bubbleEl.classList.add("hidden"), ms);
  }

  helpBtn?.addEventListener("click", ()=> helpModal?.classList.remove("hidden"));
  closeHelpBtn?.addEventListener("click", ()=> helpModal?.classList.add("hidden"));

  // If julbild.jpg missing, show fallback art
  if (wonderImg && wonderFallback){
    wonderImg.addEventListener("error", () => {
      wonderImg.classList.add("hidden");
      wonderFallback.classList.remove("hidden");
    });
  }

  function closeWonder(){
    wonder?.classList.add("hidden");
    document.body.classList.remove("win");
    document.body.classList.remove("win-shake");
    stopConfetti();
    bubble("Äntligen hemma…");
  }
  closeWonderBtn?.addEventListener("click", closeWonder);

  // ---------- Meta / Achievements / Mood ----------
  const META_KEY = "jesper_meta_v1";
  const SPEEDRUN_MS = 45000;          // confirmed
  const CONFETTI_SURVIVE_MS = 7500;
  let meta = {
    boops: 0,
    bestMs: null,
    lastRunMs: null,
    wins: 0,
    ach: { booper:false, speedrunner:false, confetti:false, door:false }
  };
  try{
    const raw = localStorage.getItem(META_KEY);
    if (raw){
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") meta = Object.assign(meta, parsed);
      if (parsed && parsed.ach) meta.ach = Object.assign(meta.ach, parsed.ach);
    }
  } catch(e){}

  let runStartMs = null;
  let runEndMs = null;

  function saveMeta(){
    try{ localStorage.setItem(META_KEY, JSON.stringify(meta)); } catch(e){}
  }

  function fmtMs(ms){
    if (ms == null || !isFinite(ms)) return "—";
    const s = ms/1000;
    const m = Math.floor(s/60);
    const ss = (s%60);
    const whole = Math.floor(ss);
    const dec = Math.floor((ss - whole)*10);
    return `${m}:${String(whole).padStart(2,"0")}.${dec}`;
  }

  function updateAchUI(){
    if (boopCountEl) boopCountEl.textContent = String(meta.boops||0);
    if (runTimeEl) runTimeEl.textContent = fmtMs(meta.lastRunMs);
    if (bestTimeEl) bestTimeEl.textContent = fmtMs(meta.bestMs);

    if (!achModal) return;
    const items = achModal.querySelectorAll(".achItem");
    items.forEach(li=>{
      const key = li.getAttribute("data-ach");
      const on = !!(meta.ach && meta.ach[key]);
      li.classList.toggle("is-on", on);
      const badge = li.querySelector(".achBadge");
      if (badge) badge.textContent = on ? "✅" : "⬜";
    });
  }

  function unlockAch(key){
    if (!meta.ach) meta.ach = {};
    if (meta.ach[key]) return;
    meta.ach[key] = true;
    saveMeta();
    updateAchUI();
    toast("🏆 Achievement unlocked!", 1400);
    sfxAchievement();
  }

  function markRunStart(){
    if (runStartMs != null || state.unlocked) return;
    runStartMs = nowMs();
  }

  function setMood(text){
    if (moodEl) moodEl.textContent = text;
  }

  function updateMood(){
    if (state.unlocked){ setMood("Mood: 🤩"); return; }
    if (state.secretStep >= 3){ setMood("Mood: 😈"); return; }
    if (state.secretStep === 2){ setMood("Mood: 😏"); return; }
    if (state.secretStep === 1){ setMood("Mood: 🤔"); return; }
    if ((state.failCount||0) >= 5){ setMood("Mood: 😵"); return; }
    setMood("Mood: 😐");
  }

  // ---------- Confetti (canvas overlay) ----------
  const confetti = {
    on:false,
    parts:[],
    raf:0,
    w:0, h:0, dpr:1,
    ctx:null,
    lastT:0,
    aliveUntil:0
  };

  function resizeConfetti(){
    if (!confettiCanvas) return;
    const dpr = Math.max(1, Math.min(2, window.devicePixelRatio||1));
    confetti.dpr = dpr;
    confetti.w = Math.floor(window.innerWidth*dpr);
    confetti.h = Math.floor(window.innerHeight*dpr);
    confettiCanvas.width = confetti.w;
    confettiCanvas.height = confetti.h;
    confettiCanvas.style.width = "100vw";
    confettiCanvas.style.height = "100vh";
    confetti.ctx = confettiCanvas.getContext("2d");
  }
  window.addEventListener("resize", resizeConfetti);
  resizeConfetti();

  function confettiBurst(normX, normY, count, power=1){
    if (!confetti.ctx) return;
    const w = confetti.w, h = confetti.h, dpr = confetti.dpr;
    const x0 = normX * w;
    const y0 = normY * h;
    const palette = ["#ef4444","#f59e0b","#eab308","#22c55e","#06b6d4","#3b82f6","#a855f7","#ec4899","#ffffff"];
    for (let i=0;i<count;i++){
      const a = (-Math.PI/2) + (Math.random()-0.5)*1.5;
      const sp = (520 + Math.random()*760) * power * dpr;
      confetti.parts.push({
        x:x0, y:y0,
        vx: Math.cos(a)*sp + (Math.random()-0.5)*120*dpr,
        vy: Math.sin(a)*sp - (Math.random()*240*dpr),
        g: (900 + Math.random()*650) * dpr,
        rot: Math.random()*Math.PI*2,
        vr: (Math.random()-0.5)*10,
        w: (6+Math.random()*9)*dpr,
        h: (3+Math.random()*7)*dpr,
        c: palette[(Math.random()*palette.length)|0],
        life: 1.1 + Math.random()*1.2
      });
    }
  }

  function startConfettiWin(){
    resizeConfetti();
    confetti.on = true;
    confetti.aliveUntil = nowMs() + 14000;
    // side cannons + center blast
    confettiBurst(0.06, 0.92, 260, 1.12);
    confettiBurst(0.94, 0.92, 260, 1.12);
    confettiBurst(0.50, 0.35, 360, 1.05);

    // periodic mini bursts
    for (let i=1;i<=6;i++){
      setTimeout(()=>{
        if (!confetti.on) return;
        confettiBurst(0.15 + Math.random()*0.70, 0.30 + Math.random()*0.45, 110, 0.75);
      }, i*650);
    }

    if (!confetti.raf){
      confetti.lastT = nowMs();
      confetti.raf = requestAnimationFrame(confettiLoop);
    }
  }

  function stopConfetti(){
    confetti.on = false;
    confetti.parts.length = 0;
    if (confetti.ctx){
      confetti.ctx.clearRect(0,0,confetti.w,confetti.h);
    }
  }

  function confettiLoop(t){
    const ctxc = confetti.ctx;
    if (!ctxc){
      confetti.raf = 0;
      return;
    }
    const now = nowMs();
    const dt = Math.min(0.033, (now - confetti.lastT)/1000);
    confetti.lastT = now;

    ctxc.clearRect(0,0,confetti.w,confetti.h);

    if (confetti.on){
      // auto stop after aliveUntil, unless wonder is still open (let it drift a bit)
      if (now > confetti.aliveUntil && (wonder?.classList.contains("hidden") ?? true)){
        stopConfetti();
      }
    }

    // update & draw
    const parts = confetti.parts;
    for (let i=parts.length-1;i>=0;i--){
      const p = parts[i];
      p.life -= dt;
      if (p.life <= 0){
        parts.splice(i,1);
        continue;
      }
      p.vy += p.g * dt;
      p.x += p.vx * dt;
      p.y += p.vy * dt;
      p.rot += p.vr * dt;

      // simple drag
      p.vx *= (1 - dt*0.22);
      p.vy *= (1 - dt*0.08);

      ctxc.save();
      ctxc.translate(p.x, p.y);
      ctxc.rotate(p.rot);
      ctxc.fillStyle = p.c;
      ctxc.globalAlpha = Math.max(0, Math.min(1, p.life));
      ctxc.fillRect(-p.w/2, -p.h/2, p.w, p.h);
      ctxc.restore();
    }

    // keep looping if any parts or active
    if (confetti.on || parts.length){
      confetti.raf = requestAnimationFrame(confettiLoop);
    } else {
      confetti.raf = 0;
    }
  }

  // ---------- Modal wiring ----------
  achBtn?.addEventListener("click", ()=>{
    updateAchUI();
    achModal?.classList.remove("hidden");
    sfxHintPop();
  });
  closeAchBtn?.addEventListener("click", ()=> achModal?.classList.add("hidden"));

  hintBtn?.addEventListener("click", ()=>{
    sfxHintPop();
    const snark = [
      "Calculated.",
      "100% intentional.",
      "Ny plan: bättre plan.",
      "Jag ser… saker… 👀"
    ];
    if (state.unlocked){
      bubble("Du vann ju?! Tryck på 🚪 Hidden door.", 1700);
      return;
    }
    if (state.secretStep === 0){
      bubble("Börja med ⏰. Jag säger inget mer.", 1700);
    } else if (state.secretStep === 1){
      bubble("Socker löser allt… 🍬", 1700);
    } else if (state.secretStep === 2){
      bubble("En ⭐ har en plan. Du också.", 1700);
    } else if (state.secretStep >= 3){
      bubble("🪑 SIT. Nu.", 1700);
    } else {
      bubble(snark[(Math.random()*snark.length)|0], 1500);
    }
  });

  doorLink?.addEventListener("click", ()=>{
    unlockAch("door");
    toast("🚪 Hidden door öppnad!", 1200);
  });

  highFiveBtn?.addEventListener("click", ()=>{
    setAction("highfive");
    setFace("ecstatic", 1.2);
    sfxSlap();
    bubble("✋ HIGH‑FIVE!", 1200);
    confettiBurst(0.50, 0.42, 220, 1.05);
  });


// ---------- Audio (WebAudio + MP3) ----------
// Tips: iOS kräver "user gesture" för att starta ljud. Därför auto-startar vi (om sparat som PÅ)
// först vid första tryck/drag på sidan.
let audioEnabled = false;
let audioCtx = null;
let bgMusic = null; // HTMLAudioElement (mp3)
// Background music should sit under SFX (lower than before)
const MUSIC_TARGET_VOL = 0.1;

// fade-state
let _fadeRaf = 0;
let _fadeToken = 0;
let pendingAutoStart = false;

function vibe(pattern){
  try {
    if (navigator && typeof navigator.vibrate === "function") navigator.vibrate(pattern);
  } catch (e) {}
}

function ensureAudio(){
  if (audioCtx) return;
  audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  try { if (audioCtx.state === "suspended" && audioCtx.resume) audioCtx.resume(); } catch (e) {}
}

function ensureMusic(){
  if (bgMusic) return;
  bgMusic = new Audio("assets/julsang.mp3");
  bgMusic.loop = true;
  bgMusic.preload = "auto";
  bgMusic.volume = MUSIC_TARGET_VOL;
}

function cancelFade(){
  if (_fadeRaf){
    try { cancelAnimationFrame(_fadeRaf); } catch (e) {}
    _fadeRaf = 0;
  }
}

function fadeVolume(audio, toVol, ms, onDone){
  cancelFade();
  const token = ++_fadeToken;
  const fromVol = clamp((audio && typeof audio.volume === "number") ? audio.volume : 0, 0.0001, 1);
  const start = nowMs();
  const dur = Math.max(1, ms|0);
  const target = clamp(toVol, 0.0001, 1);

  const step = () => {
    if (token !== _fadeToken) return;
    const t = clamp((nowMs() - start) / dur, 0, 1);
    const v = lerp(fromVol, target, t);
    try { audio.volume = v; } catch (e) {}
    if (t < 1){
      _fadeRaf = requestAnimationFrame(step);
    } else {
      _fadeRaf = 0;
      if (typeof onDone === "function") onDone();
    }
  };
  _fadeRaf = requestAnimationFrame(step);
}

async function musicOn(){
  if (!audioEnabled) return;
  ensureMusic();
  try {
    // Starta tyst och fadda in
    bgMusic.volume = 0.0001;
    await bgMusic.play();
    fadeVolume(bgMusic, MUSIC_TARGET_VOL, 650);
  } catch (e){
    console.warn("Kunde inte starta julstämningen", e);
  }
}

function musicOff(reset=true){
  if (!bgMusic) return;
  try {
    fadeVolume(bgMusic, 0.0001, 450, () => {
      try { bgMusic.pause(); } catch (e) {}
      if (reset){
        try { bgMusic.currentTime = 0; } catch (e) {}
      }
      // återställ volym så nästa start kan fadda in från tyst
      try { bgMusic.volume = MUSIC_TARGET_VOL; } catch (e) {}
    });
  } catch (e){
    // fallback: hard stop
    try { bgMusic.pause(); } catch (e2) {}
    if (reset){
      try { bgMusic.currentTime = 0; } catch (e2) {}
    }
  }
}

function setSoundBtnLabel(){
  if (!soundBtn) return;
  soundBtn.textContent = audioEnabled ? "🔊 Ljud: PÅ" : "🔊 Ljud: AV";
}

// --- restore persisted audio preference ---
try {
  audioEnabled = localStorage.getItem("jesper_audio_enabled") === "1";
  pendingAutoStart = audioEnabled;
} catch (e) {}
setSoundBtnLabel();
updateAchUI();
if (audioEnabled) setTimeout(()=>toast("Ljud sparat som PÅ – rör skärmen för att starta.", 2200), 650);

async function startAudioFromGesture(){
  if (!audioEnabled || !pendingAutoStart) return;
  pendingAutoStart = false;
  ensureAudio();
  if (audioCtx && audioCtx.state === "suspended"){
    try { await audioCtx.resume(); } catch (e) {}
  }
  await musicOn();
}

// Om användaren tidigare haft ljud PÅ: starta vid första interaktion (drag på canvas räcker).
window.addEventListener("pointerdown", startAudioFromGesture, { once:true, passive:true });
window.addEventListener("touchstart", startAudioFromGesture, { once:true, passive:true });
window.addEventListener("mousedown", startAudioFromGesture, { once:true, passive:true });

function ping(freq=440, t=0.08, gain=0.09, type="triangle"){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  const o = a.createOscillator();
  const g = a.createGain();
  o.type = type;
  o.frequency.setValueAtTime(freq, now);
  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(gain, now + 0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, now + t);

  o.connect(g); g.connect(a.destination);
  o.start(now); o.stop(now + t + 0.02);
}

function grunt(intensity=1){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  // thump
  const o = a.createOscillator();
  const g = a.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(140, now);
  o.frequency.exponentialRampToValueAtTime(52, now + 0.12);

  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(0.28*intensity, now + 0.006);
  g.gain.exponentialRampToValueAtTime(0.0001, now + 0.16);

  o.connect(g); g.connect(a.destination);
  o.start(now); o.stop(now + 0.18);

  // grit
  ping(260 + Math.random()*80, 0.07, 0.08*intensity, "sawtooth");
}

function pop(intensity=1){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  // tiny "spark" + noise puff
  const o = a.createOscillator();
  const g = a.createGain();
  o.type = "square";
  o.frequency.setValueAtTime(700 + Math.random()*500*intensity, now);
  o.frequency.exponentialRampToValueAtTime(220, now + 0.09);

  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(0.12*intensity, now + 0.005);
  g.gain.exponentialRampToValueAtTime(0.0001, now + 0.11);

  o.connect(g); g.connect(a.destination);
  o.start(now); o.stop(now + 0.13);

  // click
  ping(1200 + Math.random()*500, 0.035, 0.05*intensity, "triangle");
}

function swoosh(){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  // whoosh via noise
  const bufferSize = a.sampleRate * 0.18;
  const buffer = a.createBuffer(1, bufferSize, a.sampleRate);
  const data = buffer.getChannelData(0);
  for (let i=0;i<data.length;i++){
    data[i] = (Math.random()*2-1) * Math.pow(1 - i/data.length, 1.7);
  }
  const noise = a.createBufferSource();
  noise.buffer = buffer;

  const bp = a.createBiquadFilter();
  bp.type = "bandpass";
  bp.frequency.setValueAtTime(900, now);
  bp.frequency.exponentialRampToValueAtTime(320, now + 0.18);
  bp.Q.value = 0.8;

  const g = a.createGain();
  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(0.22, now + 0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, now + 0.20);

  noise.connect(bp); bp.connect(g); g.connect(a.destination);
  noise.start(now); noise.stop(now + 0.22);
}

function jingle(){
  if (!audioEnabled) return;
  // victory-ish riff
  const base = 523.25;
  const notes = [1, 5/4, 3/2, 2, 3/2, 5/4, 1];
  notes.forEach((n,i)=>{
    setTimeout(()=>ping(base*n, 0.09, 0.10), i*95);
  });
  setTimeout(()=>ping(base*2.5, 0.12, 0.07, "sine"), 580);
}

function sfxPunchHit(){
  swoosh();
  grunt(1.1);
  // sparkle on hit
  ping(660 + Math.random()*60, 0.06, 0.06, "triangle");
}

function sfxPunchMiss(){
  swoosh();
  ping(180 + Math.random()*40, 0.06, 0.06, "sawtooth");
}

function sfxNope(){
  ping(210, 0.08, 0.08, "square");
  ping(140, 0.10, 0.07, "sawtooth");
}

function sfxNopeSoft(){
  ping(240, 0.06, 0.045, "triangle");
}

function sfxStep(step, combo){
  // rising "almost solved" tones
  const base = 392; // G4
  const up = [0, 0, 2, 4][clamp(step,0,3)|0] || 0;
  const c = clamp(combo||1,1,4);
  const gain = 0.06 + (c-1)*0.015;
  ping(base * Math.pow(2, up/12), 0.09, gain, "triangle");
  setTimeout(()=>ping(base * Math.pow(2, (up+7)/12), 0.07, gain*0.8, "triangle"), 70);
  if (step >= 3) setTimeout(()=>ping(base*2, 0.10, gain*0.9, "sine"), 150);
}

function sfxTension(){
  // tiny tension tick
  ping(320, 0.06, 0.05, "sawtooth");
  setTimeout(()=>ping(420, 0.06, 0.05, "sawtooth"), 80);
}

function sfxHintPop(){
  pop(0.9);
}

function sfxAchievement(){
  // sparkle ladder
  [880, 1046, 1318].forEach((f,i)=>setTimeout(()=>ping(f, 0.06, 0.07, "triangle"), i*60));
}

function sfxBassDrop(){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  const o = a.createOscillator();
  const g = a.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(140, now);
  o.frequency.exponentialRampToValueAtTime(28, now + 0.52);
  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(0.35, now + 0.02);
  g.gain.exponentialRampToValueAtTime(0.0001, now + 0.60);

  o.connect(g); g.connect(a.destination);
  o.start(now); o.stop(now + 0.62);
}

function sfxCrowdWoo(){
  if (!audioEnabled) return;
  ensureAudio();
  const a = audioCtx;
  const now = a.currentTime;

  // breathy "woo" approximation (noise through sweeping bandpass)
  const bufferSize = a.sampleRate * 0.55;
  const buffer = a.createBuffer(1, bufferSize, a.sampleRate);
  const data = buffer.getChannelData(0);
  for (let i=0;i<data.length;i++){
    const t = i/data.length;
    data[i] = (Math.random()*2-1) * (1 - t) * 0.7;
  }
  const n = a.createBufferSource();
  n.buffer = buffer;

  const bp = a.createBiquadFilter();
  bp.type = "bandpass";
  bp.frequency.setValueAtTime(380, now);
  bp.frequency.exponentialRampToValueAtTime(720, now + 0.28);
  bp.frequency.exponentialRampToValueAtTime(420, now + 0.55);
  bp.Q.value = 0.9;

  const g = a.createGain();
  g.gain.setValueAtTime(0.0001, now);
  g.gain.exponentialRampToValueAtTime(0.22, now + 0.02);
  g.gain.exponentialRampToValueAtTime(0.0001, now + 0.55);

  n.connect(bp); bp.connect(g); g.connect(a.destination);
  n.start(now); n.stop(now + 0.58);
}

function sfxSlap(){
  pop(1.05);
  ping(980, 0.04, 0.08, "square");
}

function sfxBoop(){
  ping(660 + Math.random()*80, 0.05, 0.06, "triangle");
  setTimeout(()=>ping(990 + Math.random()*90, 0.05, 0.05, "triangle"), 70);
}

function sfxSit(){
  // chair creak
  ping(190 + Math.random()*25, 0.10, 0.06, "sawtooth");
  setTimeout(()=>ping(120 + Math.random()*20, 0.08, 0.05, "triangle"), 90);
}

function sfxDance(){
  // tiny funky 3-hit
  const b = 440;
  [0, 4, 7].forEach((st,i)=>setTimeout(()=>ping(b*Math.pow(2, st/12), 0.06, 0.07, "triangle"), i*90));
}

function sfxJuggle(){
  ping(520, 0.06, 0.05, "triangle");
  setTimeout(()=>ping(640, 0.06, 0.05, "triangle"), 80);
}

function sfxDrop(){
  ping(160, 0.07, 0.08, "square");
  setTimeout(()=>ping(110, 0.08, 0.07, "sawtooth"), 70);
}



soundBtn?.addEventListener("click", async ()=>{
  audioEnabled = !audioEnabled;
  pendingAutoStart = false;

  try { localStorage.setItem("jesper_audio_enabled", audioEnabled ? "1" : "0"); } catch (e) {}
  setSoundBtnLabel();

  if (audioEnabled){
    ensureAudio();
    if (audioCtx && audioCtx.state === "suspended") await audioCtx.resume();
    await musicOn();
    toast("Ljud på.");
    grunt(1.0);
    vibe(18);
  } else {
    musicOff(true);
    toast("Ljud av.");
    vibe(10);
  }
});

// Om användaren byter app/flik: pausa musiken så iOS inte blir grinig.
  document.addEventListener("visibilitychange", ()=>{
    if (!bgMusic) return;
    if (document.hidden){
      musicOff(false);
    } else if (audioEnabled){
      musicOn();
    }
  }, { passive:true });

  // ---------- Canvas / rendering ----------
  if (!canvas){
    console.warn("Canvas #game saknas.");
    return;
  }
  const ctx = canvas.getContext("2d", { alpha: false });

  // fixed world (side view)
  const WORLD = { w: 900, h: 360 };
  const view = { s: 1, ox: 0, oy: 0, cssW: 0, cssH: 0 };

  // room rectangle (side view)
  const ROOM = { x: 40, y: 40, w: 820, h: 260 };
  const FLOOR_Y = ROOM.y + ROOM.h - 38; // ground line

  function resize(){
    const rect = canvas.getBoundingClientRect();
    const dpr = Math.min(window.devicePixelRatio || 1, 2);

    canvas.width = Math.max(1, Math.round(rect.width * dpr));
    canvas.height = Math.max(1, Math.round(rect.height * dpr));
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

    view.cssW = rect.width;
    view.cssH = rect.height;

    view.s = Math.min(rect.width / WORLD.w, rect.height / WORLD.h);
    view.ox = (rect.width - WORLD.w * view.s) / 2;
    view.oy = (rect.height - WORLD.h * view.s) / 2;
  }
  window.addEventListener("resize", resize, { passive:true });
  resize();

  function toWorld(px, py){
    return { x: (px - view.ox)/view.s, y: (py - view.oy)/view.s };
  }

  // ---------- Game objects ----------
  const props = {
    table: { x: 150, w: 210 },
    chair: { x: 430, w: 130 },
    tree:  { x: 700, w: 130 },
    frame: { x: 90, y: 75, w: 90, h: 60 } // empty “tavla”
  };

  const jesper = {
    x: 120,
    vx: 0,
    facing: 1,
    r: 22,
    action: "idle", // idle/walk/kick/sit/bump/wave
    actionT: 0,
    blinkT: 0,
    idleTimer: 0,
    bumpCd: 0,
    faceMood: 'neutral',
    faceT: 0,
    browT: 0,
    tilt: 0,

  };

  const ornaments = [
    { id:"clock", label:"⏰", x: 300, vx: 0, r: 20, base:"#fde047" },
    { id:"candy", label:"🍬", x: 360, vx: 0, r: 20, base:"#fb7185" },
    { id:"star",  label:"⭐", x: 520, vx: 0, r: 20, base:"#60a5fa" }
  ];

  const state = {
    joy: { active:false, startX:0, dx:0 },
    dragging: null,
    secretStep: 0,
    unlocked: false,
    shakeT: 0,
    shakeMag: 0,
    failCount: 0,
    combo: 0,
    lastCorrectMs: 0,
    pointer: null,
    boopCd: 0,
    timeScale: 1,
    slowMoT: 0,
    frameFx: { glitchT:0, freezeT:0, winkT:0, nextAt:0, freezeCanvas:null }

  };

  updateMood();
  updateAchUI();


  // ---------- Persist unlock (localStorage) ----------
  try {
    state.unlocked = localStorage.getItem("jesper_unlocked") === "1";
    if (state.unlocked){
    setTimeout(()=>wonder?.classList.remove("hidden"), 800);
    danceBtn?.classList.remove("hidden");
    updateMood();
  }
  } catch (e) {}

  // keep everything on floor
  function floorYForRadius(r){ return FLOOR_Y - r; }

  // collision x ranges for ornaments
  function blockRanges(){
    // each block is a solid column on floor (table legs-ish, chair base, tree trunk)
    return [
      { x: props.table.x + 18, w: 28 },
      { x: props.table.x + props.table.w - 46, w: 28 },
      { x: props.chair.x + 12, w: props.chair.w - 24 },
      { x: props.tree.x + 52, w: 26 }, // trunk
    ];
  }

  function resolveOrnamentBlocks(o){
    // room bounds
    const minX = ROOM.x + o.r;
    const maxX = ROOM.x + ROOM.w - o.r;
    if (o.x < minX){ o.x = minX; o.vx *= -0.55; }
    if (o.x > maxX){ o.x = maxX; o.vx *= -0.55; }

    // blocks
    for (const b of blockRanges()){
      const left = b.x - o.r;
      const right = b.x + b.w + o.r;
      if (o.x > left && o.x < right){
        // push out to nearest side
        const dl = Math.abs(o.x - left);
        const dr = Math.abs(right - o.x);
        if (dl < dr){
          o.x = left;
          o.vx = -Math.abs(o.vx) * 0.65;
        } else {
          o.x = right;
          o.vx = Math.abs(o.vx) * 0.65;
        }
      }
    }
  }

  function setAction(name){
    jesper.action = name;
    jesper.actionT = 0;
  }

  function setFace(mode, seconds=0.9){
    jesper.faceMood = mode || "neutral";
    jesper.faceT = Math.max(0, seconds);
  }

  function faceTick(dt){
    jesper.faceT = Math.max(0, (jesper.faceT||0) - dt);
    if ((jesper.faceT||0) <= 0){
      jesper.faceMood = "neutral";
    }
    jesper.browT = Math.max(0, (jesper.browT||0) - dt);
  }


  function onWallBump(side){
    if (jesper.bumpCd > 0) return;
    jesper.bumpCd = 0.55;
    jesper.facing = -jesper.facing;
    setAction("bump");
    grunt(0.65);
    vibe(14);
    bubble("💢 Aj!", 900);
    pop(0.9);
    // liten skakning
    state.shakeT = 0.18;
    state.shakeMag = 7;
  }


  // ---------- Input (pointer) ----------
  function pointerPos(e){
    const r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  canvas.addEventListener("pointerdown", (e)=>{
    e.preventDefault();
    const p = pointerPos(e);
    const w = toWorld(p.x, p.y);
    state.pointer = w;
    markRunStart();

    // secret boop pixel (near frame)
    if ((state.boopCd||0) <= 0){
      const f = props.frame;
      const bx = f.x + f.w - 10;
      const by = f.y + f.h - 10;
      if (w.x >= bx && w.x <= bx + 6 && w.y >= by && w.y <= by + 6){
        state.boopCd = 0.25;
        meta.boops = (meta.boops||0) + 1;
        meta.lastRunMs = meta.lastRunMs; // no-op
        saveMeta();
        updateAchUI();
        toast(`boop +1 (tot: ${meta.boops})`, 1100);
        sfxBoop();
        if (meta.boops === 1) unlockAch("booper");
        return;
      }
    }

    // hit ornament? (drag)
    for (const o of ornaments){
      const oy = floorYForRadius(o.r);
      const dx = w.x - o.x;
      const dy = w.y - oy;
      if (Math.hypot(dx,dy) <= o.r + 10){
        state.dragging = o;
        o.vx = 0;
        canvas.setPointerCapture(e.pointerId);
        toast("Flyttar pynt.");
        grunt(0.7);
        return;
      }
    }

    // else joystick (horizontal only)
    state.joy.active = true;
    state.joy.startX = w.x;
    state.joy.dx = 0;
    canvas.setPointerCapture(e.pointerId);
  }, { passive:false });

  canvas.addEventListener("pointermove", (e)=>{
    e.preventDefault();
    const p = pointerPos(e);
    const w = toWorld(p.x, p.y);
    state.pointer = w;

    if (state.dragging){
      state.dragging.x = clamp(w.x, ROOM.x + state.dragging.r, ROOM.x + ROOM.w - state.dragging.r);
      return;
    }

    if (state.joy.active){
      state.joy.dx = clamp(w.x - state.joy.startX, -120, 120);
    }
  }, { passive:false });

  function endPointer(){
    state.dragging = null;
    state.joy.active = false;
    state.joy.dx = 0;
  }
  canvas.addEventListener("pointerup", endPointer, { passive:false });
  canvas.addEventListener("pointercancel", endPointer, { passive:false });


// Touch fallback (äldre iOS/Android): joystick på canvas (horisontell)
if (!("PointerEvent" in window)){
  canvas.addEventListener("touchstart", (e)=>{
    if (e.touches.length !== 1) return;
    e.preventDefault();
    const t = e.touches[0];
    const r = canvas.getBoundingClientRect();
    const x = t.clientX - r.left;
    const y = t.clientY - r.top;
    const w = toWorld(x, y);
    state.joy.active = true;
    state.joy.startX = w.x;
    state.joy.dx = 0;
  }, { passive:false });

  canvas.addEventListener("touchmove", (e)=>{
    if (!state.joy.active || e.touches.length !== 1) return;
    e.preventDefault();
    const t = e.touches[0];
    const r = canvas.getBoundingClientRect();
    const x = t.clientX - r.left;
    const y = t.clientY - r.top;
    const w = toWorld(x, y);
    state.joy.dx = clamp(w.x - state.joy.startX, -120, 120);
  }, { passive:false });

  canvas.addEventListener("touchend", ()=>endPointer(), { passive:true });
  canvas.addEventListener("touchcancel", ()=>endPointer(), { passive:true });
}

  // ---------- Buttons ----------
  kickBtn?.addEventListener("click", ()=>doKick());
  sitBtn?.addEventListener("click", ()=>doSit());
  danceBtn?.addEventListener("click", ()=>doDance());

  // ---------- Secret logic ----------
  function advanceSecret(id){
    if (state.unlocked) return;
    const seq = ["clock","candy","star"];
    const expected = seq[state.secretStep] || seq[0];

    // wrong -> reset
    if (id !== expected){
      state.secretStep = 0;
      state.combo = 0;
      state.failCount = (state.failCount||0) + 1;
      toast("Nää… Börja om…", 1200);
      setFace("frown", 0.85);
      sfxNope();

      if (state.failCount % 5 === 0){
        const taunts = [
          "Vi låtsas att det där inte hände.",
          "Ny strategi: sikta bättre 😇",
          "Det där var… modigt.",
          "Jag ser vad du försöker. Det är fel.",
          "Jag tror på dig. Typ."
        ];
        bubble(taunts[(Math.random()*taunts.length)|0], 1700);
      }
      updateMood();
      return;
    }

    // correct
    markRunStart();
    state.secretStep++;
    state.failCount = 0;

    // combo tracking
    const n = nowMs();
    const last = state.lastCorrectMs || 0;
    state.combo = (n - last) < 2200 ? (state.combo||0) + 1 : 1;
    state.lastCorrectMs = n;

    toast(`Hemligheten: ${state.secretStep}/3`, 1100);

    // escalating chime + mood
    sfxStep(state.secretStep, state.combo);

    if (state.combo >= 3) setFace("star", 1.25);
    else if (state.combo === 2) setFace("grin", 1.10);
    else setFace("smile", 0.95);

    // overconfident one-liners sometimes
    if (Math.random() < 0.25){
      const lines = ["Calculated.", "100% intentional.", "Jag är så smart.", "Lätt."];
      bubble(lines[(Math.random()*lines.length)|0], 1200);
    }

    if (state.secretStep === 3){
      bubble("SITT på stolen. Nu.", 1500);
      sfxTension();
      // increase frame shenanigans when close
      state.frameFx.nextAt = 0;
    }

    updateMood();
  }

  function unlock(){
    if (state.unlocked) return;
    state.unlocked = true;
    try { localStorage.setItem("jesper_unlocked", "1"); } catch (e) {}

    // freeze input
    state.joy.active = false;
    state.dragging = null;

    // run time / achievements
    runEndMs = nowMs();
    const runMs = (runStartMs != null) ? Math.max(0, runEndMs - runStartMs) : null;
    meta.lastRunMs = runMs;
    meta.wins = (meta.wins||0) + 1;
    if (runMs != null){
      if (meta.bestMs == null || runMs < meta.bestMs) meta.bestMs = runMs;
      if (runMs <= SPEEDRUN_MS) unlockAch("speedrunner");
    }
    saveMeta();
    updateAchUI();

    state.slowMoT = 0.9;

    // win visuals
    document.body.classList.add("win");
    document.body.classList.add("win-shake");
    setTimeout(()=>document.body.classList.remove("win-shake"), 850);

    startConfettiWin();

    // win audio stack
    sfxBassDrop();
    setTimeout(()=>{ jingle(); sfxCrowdWoo(); }, 520);

    vibe([30,40,30,60,30]);
    setFace("ecstatic", 2.2);
    bubble("…okej. Du vann. 😤✨", 1700);

    // show extra action button
    danceBtn?.classList.remove("hidden");

    wonder?.classList.remove("hidden");
    updateMood();

    // confetti survivor: keep the modal open for a bit
    setTimeout(()=>{
      if (state.unlocked && !(wonder?.classList.contains("hidden") ?? true) && confetti.on){
        unlockAch("confetti");
      }
    }, CONFETTI_SURVIVE_MS);
  }

  // ---------- Actions ----------
  function doKick(){
    markRunStart();
    const reach = 90;

    setAction("kick");
    // miss/hit SFX differs after we know hit
    vibe(22);

    let best = null, bestD = 1e9;
    for (const o of ornaments){
      const d = Math.abs(o.x - jesper.x);
      if (d < bestD){
        bestD = d; best = o;
      }
    }

    if (!best || bestD > reach){
      sfxPunchMiss();
      setFace("neutral", 0.6);
      bubble("Miss! Men det räknas.", 1200);
      return;
    }

    // HIT
    const dir = Math.sign(best.x - jesper.x) || jesper.facing;
    best.vx += dir * (520 + Math.random()*120);

    sfxPunchHit();
    setFace("smile", 1.0);
    bubble(`👊 HIT! (${best.label})`, 900);

    // micro: eyebrow pop when you do the correct step
    jesper.browT = 0.6;

    advanceSecret(best.id);
  }

  function doSit(){
    markRunStart();
    setAction("sit");
    sfxSit();
    vibe(12);

    if (state.unlocked){
      bubble("🪑 …existens… jul…", 1400);
      return;
    }

    if (state.secretStep === 3 && !state.unlocked){
      toast("Uppdrag utfört!", 1400);
      unlock();
    } else {
      bubble("🪑 Skönt. Men… inte än.", 1400);
      setFace("neutral", 0.9);
      sfxNopeSoft();
    }
  }


  function doDance(){
    if (!state.unlocked){
      bubble("Nej. Förtjäna dansen först.", 1400);
      sfxNopeSoft();
      return;
    }
    markRunStart();
    setAction("dance");
    setFace("grin", 1.6);
    sfxDance();
    confettiBurst(0.50, 0.55, 140, 0.75);
    bubble("🕺 JESP-DANS! (ingen filmning)", 1600);
  }



  // ---------- Commentary ----------
  const lines = [
    "Det här rummet känns… budget…",
    "Kanske en kaffe…",
    "Jag är 100% ledig…",
    "Den där lilla tavlan…",
    "Undra hur det går på fabriken…"
  ];
  setInterval(()=>{
    if (!wonder?.classList.contains("hidden")) return;
    if (Math.random() < 0.22){
      bubble(lines[(Math.random()*lines.length)|0], 1700);
      grunt(0.5);
    }
  }, 4200);

  // ---------- Update / Draw ----------
  function update(dt, tMs){
    // paused? (wonder open)
    const paused = !(wonder?.classList.contains("hidden") ?? true);

    // slow-mo decay (win moment)
    state.slowMoT = Math.max(0, (state.slowMoT||0) - dt);
    state.timeScale = (state.slowMoT > 0) ? 0.22 : 1.0;
    dt = dt * state.timeScale;

    // blink timer
    jesper.blinkT -= dt;
    if (jesper.blinkT <= 0) jesper.blinkT = 2.5 + Math.random()*2.2;

    // face timers
    faceTick(dt);

    // boop cooldown
    state.boopCd = Math.max(0, (state.boopCd||0) - dt);

    // frame FX timing (glitch / freeze / wink)
    const fx = state.frameFx;
    if (fx){
      fx.glitchT = Math.max(0, (fx.glitchT||0) - dt);
      fx.freezeT = Math.max(0, (fx.freezeT||0) - dt);
      fx.winkT   = Math.max(0, (fx.winkT||0) - dt);

      const spicy = clamp(state.secretStep / 3, 0, 1);
      const baseGap = 4200 - spicy*2200; // 4.2s -> 2.0s
      if (!state.unlocked && (state.secretStep >= 1) && (tMs > (fx.nextAt||0))){
        const r = Math.random();
        if (r < 0.55){
          fx.glitchT = 0.18 + Math.random()*0.22;
        } else {
          fx.freezeT = 0.32 + Math.random()*0.22;
          fx.winkT = 0.22 + Math.random()*0.18;
          fx.freezeCanvas = null; // will capture on draw
        }
        fx.nextAt = tMs + baseGap*(0.75 + Math.random()*0.65);
      }
    }

    // If paused: keep tiny idle animation only
    if (paused){
      // tiny bobbing idle
      jesper.vx = lerp(jesper.vx, 0, clamp(dt*10,0,1));
      jesper.actionT += dt;
      if (jesper.action === "dance" && jesper.actionT > 1.8) jesper.action = "idle";
      if (jesper.action === "highfive" && jesper.actionT > 0.7) jesper.action = "idle";
      return;
    }

    // action state timing
    jesper.actionT += dt;
    if (jesper.action === "kick" && jesper.actionT > 0.35) jesper.action = "idle";
    if (jesper.action === "sit"  && jesper.actionT > 0.9)  jesper.action = "idle";
    if (jesper.action === "bump" && jesper.actionT > 0.40) jesper.action = "idle";
    if (jesper.action === "wave" && jesper.actionT > 1.20) jesper.action = "idle";
    if (jesper.action === "dance" && jesper.actionT > 1.8) jesper.action = "idle";
    if (jesper.action === "juggle" && jesper.actionT > 2.6) jesper.action = "idle";
    if (jesper.action === "drop" && jesper.actionT > 0.8) jesper.action = "idle";

    // movement (only x)
    let targetV = 0;
    if (state.joy.active){
      const n = state.joy.dx / 120; // -1..1
      targetV = n * 230;
    }

    if (Math.abs(targetV) > 12){
      jesper.facing = Math.sign(targetV);
      if (jesper.action !== "kick" && jesper.action !== "sit" && jesper.action !== "bump" && jesper.action !== "wave" && jesper.action !== "dance" && jesper.action !== "juggle" && jesper.action !== "drop") jesper.action = "walk";
    } else {
      if (jesper.action === "walk") jesper.action = "idle";
    }

    // smooth velocity
    jesper.vx = lerp(jesper.vx, targetV, clamp(dt*12, 0, 1));

    jesper.x += jesper.vx * dt;
    const leftWall  = ROOM.x + jesper.r;
    const rightWall = ROOM.x + ROOM.w - jesper.r;

    // wall bump
    jesper.bumpCd = Math.max(0, (jesper.bumpCd||0) - dt);
    if (jesper.x < leftWall){
      jesper.x = leftWall;
      if (jesper.vx < -40) onWallBump(-1);
      jesper.vx = Math.max(0, -jesper.vx * 0.35);
    } else if (jesper.x > rightWall){
      jesper.x = rightWall;
      if (jesper.vx > 40) onWallBump(1);
      jesper.vx = Math.min(0, -jesper.vx * 0.35);
    }

    // head tilt based on pointer proximity
    if (state.pointer){
      const dx = state.pointer.x - jesper.x;
      jesper.tilt = lerp(jesper.tilt||0, clamp(dx/220, -1, 1), clamp(dt*6,0,1));
      // eyebrow raise if you hover near the WRONG next thing
      if (!state.unlocked){
        const seq = ["clock","candy","star"];
        const expected = seq[state.secretStep] || "clock";
        let nearest = null, nearestD = 1e9;
        for (const o of ornaments){
          const d = Math.abs(o.x - state.pointer.x);
          if (d < nearestD){ nearestD = d; nearest = o; }
        }
        if (nearest && nearestD < 55 && nearest.id !== expected){
          jesper.browT = Math.max(jesper.browT||0, 0.35);
        }
      }
    } else {
      jesper.tilt = lerp(jesper.tilt||0, 0, clamp(dt*4,0,1));
    }

    // Idle trigger (efter 10s) -> wave eller juggle (och ibland drop-gag)
    jesper.idleTimer = (jesper.idleTimer || 0) + dt;
    if (jesper.action === "idle" && jesper.idleTimer > 10){
      if (Math.random() < 0.55){
        bubble("👋 Hoho...?");
        setAction("wave");
        sfxHintPop();
      } else {
        bubble("🤹 …jag kan detta…", 1300);
        setAction("juggle");
        sfxJuggle();
        setTimeout(()=>{ if (jesper.action==="juggle"){ setAction("drop"); sfxDrop(); bubble("…oops.", 1200); } }, 2200);
      }
      jesper.idleTimer = 0;
    }
    if (jesper.action !== "idle") jesper.idleTimer = 0;

    // ornaments physics (1D)
    const friction = Math.pow(0.07, dt); // strong damping
    for (const o of ornaments){
      if (state.dragging === o) continue;
      o.x += o.vx * dt;
      o.vx *= friction;
      if (Math.abs(o.vx) < 3) o.vx = 0;
      resolveOrnamentBlocks(o);
    }
  }

  function draw(tMs){
    // clear
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0,0,view.cssW, view.cssH);

    // world transform
    ctx.save();
    ctx.translate(view.ox, view.oy);
    if (state.shakeT > 0){
      state.shakeT = Math.max(0, state.shakeT - (1/60));
      const m = state.shakeMag || 0;
      ctx.translate((Math.random()-0.5)*m, (Math.random()-0.5)*m*0.6);
    }
    ctx.scale(view.s, view.s);

    // room box
    ctx.fillStyle = "#f3f4f6";
    ctx.fillRect(ROOM.x, ROOM.y, ROOM.w, ROOM.h);

    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    roundRectStroke(ROOM.x, ROOM.y, ROOM.w, ROOM.h, 18);

    // wall/floor separation
    ctx.strokeStyle = "rgba(17,24,39,0.25)";
    ctx.lineWidth = 3;
    ctx.beginPath();
    ctx.moveTo(ROOM.x, FLOOR_Y);
    ctx.lineTo(ROOM.x + ROOM.w, FLOOR_Y);
    ctx.stroke();

    // tavla (bilden styrs av <img id="tavlaImg" ...> i index.html)
    drawFrame();

    // furniture
    drawTable();
    drawChair();
    drawTree(tMs);

    // ornaments
    for (const o of ornaments) drawOrnament(o);

    // Jesper
    drawJesper(tMs);

    // joystick hint
    if (state.joy.active && !state.dragging) drawJoystick();

    // label
    ctx.fillStyle = "rgba(17,24,39,0.25)";
    ctx.font = "900 14px ui-monospace, monospace";
    ctx.fillText("                    JUL", ROOM.x + 14, ROOM.y + ROOM.h - 12);

    ctx.restore();
  }

  function roundRectStroke(x,y,w,h,r){
    const rr = Math.min(r, w/2, h/2);
    ctx.beginPath();
    ctx.moveTo(x+rr, y);
    ctx.arcTo(x+w, y, x+w, y+h, rr);
    ctx.arcTo(x+w, y+h, x, y+h, rr);
    ctx.arcTo(x, y+h, x, y, rr);
    ctx.arcTo(x, y, x+w, y, rr);
    ctx.closePath();
    ctx.stroke();
  }

  
function drawImageContain(img, x, y, w, h){
  // contain-fit (passa in) — ingen beskärning
  const iw = img.naturalWidth || img.width;
  const ih = img.naturalHeight || img.height;
  if (!iw || !ih) return false;
  const s = Math.min(w / iw, h / ih);
  const dw = iw * s;
  const dh = ih * s;
  const dx = x + (w - dw) / 2;
  const dy = y + (h - dh) / 2;
  ctx.drawImage(img, dx, dy, dw, dh);
  return true;
}

function drawFrame(){
  const f = props.frame;
  const innerX = f.x + 6, innerY = f.y + 6;
  const innerW = f.w - 12, innerH = f.h - 12;

  // ram
  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#111827";
  ctx.lineWidth = 4;
  roundRectStroke(f.x, f.y, f.w, f.h, 10);
  ctx.fillRect(f.x+4, f.y+4, f.w-8, f.h-8);

  // matt bakgrund
  ctx.fillStyle = "#f3f4f6";
  ctx.fillRect(innerX, innerY, innerW, innerH);

  let drew = false;
  const img = tavlaImg;

  const fx = state.frameFx || {glitchT:0, freezeT:0, winkT:0};
  const wantFreeze = (fx.freezeT||0) > 0;

  // capture freeze snapshot on-demand
  if (wantFreeze && !fx.freezeCanvas){
    try{
      const c = document.createElement("canvas");
      c.width = Math.max(1, Math.floor(innerW * (view.s||1)));
      c.height = Math.max(1, Math.floor(innerH * (view.s||1)));
      const cctx = c.getContext("2d");
      // draw current GIF frame into snapshot
      if (img && img.complete && img.naturalWidth){
        // draw contain into snapshot coords
        const iw = img.naturalWidth, ih = img.naturalHeight;
        const sc = Math.min(c.width/iw, c.height/ih);
        const dw = iw*sc, dh = ih*sc;
        const dx = (c.width - dw)/2, dy = (c.height - dh)/2;
        cctx.fillStyle = "#f3f4f6";
        cctx.fillRect(0,0,c.width,c.height);
        cctx.drawImage(img, dx, dy, dw, dh);
      }
      fx.freezeCanvas = c;
    }catch(e){}
  }

  if (wantFreeze && fx.freezeCanvas){
    try{
      ctx.save();
      ctx.globalAlpha = 0.98;
      // draw snapshot
      ctx.drawImage(fx.freezeCanvas, innerX, innerY, innerW, innerH);
      // tiny film grain
      ctx.globalAlpha = 0.08;
      ctx.fillStyle = "#111827";
      for (let i=0;i<14;i++){
        ctx.fillRect(innerX + Math.random()*innerW, innerY + Math.random()*innerH, 1.2, 1.2);
      }
      ctx.restore();
      drew = true;
    }catch(e){}
  } else if (img && img.complete && img.naturalWidth){
    // glitch mode
    if ((fx.glitchT||0) > 0){
      ctx.save();
      ctx.globalAlpha = 0.98;
      const slices = 6;
      for (let i=0;i<slices;i++){
        const sy = innerY + (innerH/slices)*i;
        const sh = innerH/slices;
        const jx = (Math.random()-0.5)*8;
        ctx.drawImage(img, innerX, sy, innerW, sh, innerX + jx, sy, innerW, sh);
      }
      // scanline
      ctx.globalAlpha = 0.12;
      ctx.fillStyle = "#111827";
      ctx.fillRect(innerX, innerY + Math.random()*innerH, innerW, 2);
      ctx.restore();
      drew = true;
    } else {
      drew = drawImageContain(img, innerX, innerY, innerW, innerH);
    }
  }

  if (!drew){
    drawFramePlaceholder();
  }

  // wink overlay
  if ((fx.winkT||0) > 0){
    ctx.save();
    ctx.translate(innerX + innerW*0.62, innerY + innerH*0.52);
    ctx.strokeStyle = "rgba(17,24,39,0.75)";
    ctx.lineWidth = 2.6;
    ctx.beginPath();
    ctx.moveTo(-10, 0); ctx.lineTo(-2, 0);
    ctx.moveTo( 4, 0);
    ctx.quadraticCurveTo(10, -2, 14, 0);
    ctx.stroke();
    ctx.restore();
  }

  ctx.restore();
}

function drawFramePlaceholder(){
    const f = props.frame;
    ctx.save();
    ctx.fillStyle = "#ffffff";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    roundRectStroke(f.x, f.y, f.w, f.h, 10);
    ctx.fillRect(f.x+4, f.y+4, f.w-8, f.h-8);

    ctx.fillStyle = "rgba(17,24,39,0.35)";
    ctx.font = "900 12px ui-monospace, monospace";
    ctx.fillText("TAVLA", f.x + 16, f.y + 24);
    ctx.fillStyle = "rgba(17,24,39,0.25)";
    ctx.font = "900 10px ui-monospace, monospace";
    ctx.fillText("(lägg bild själv)", f.x + 10, f.y + 42);
    ctx.restore();
  }

  function drawTable(){
    const x = props.table.x, w = props.table.w;
    const topY = FLOOR_Y - 82;
    const h = 34;

    // shadow
    ctx.save();
    ctx.fillStyle = "rgba(17,24,39,0.16)";
    ctx.beginPath();
    ctx.ellipse(x + w/2, FLOOR_Y + 2, w*0.46, 10, 0, 0, Math.PI*2);
    ctx.fill();
    ctx.restore();

    // tabletop (trä)
    ctx.fillStyle = "#f59e0b";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    roundRectStroke(x, topY, w, h, 12);
    ctx.fillRect(x + 4, topY + 4, w - 8, h - 8);

    // wood grain
    ctx.save();
    ctx.globalAlpha = 0.18;
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 1.5;
    for (let i=0;i<6;i++){
      const yy = topY + 8 + i*4.6;
      ctx.beginPath();
      ctx.moveTo(x + 10, yy);
      ctx.lineTo(x + w - 10, yy + (i%2?1.2:-0.8));
      ctx.stroke();
    }
    ctx.restore();

    // christmas table runner
    ctx.fillStyle = "#ef4444";
    ctx.fillRect(x + 14, topY + 6, w - 28, 8);
    ctx.fillStyle = "#22c55e";
    ctx.fillRect(x + 14, topY + 16, w - 28, 3);

    // legs (4)
    ctx.fillStyle = "#9ca3af";
    const legY = topY + h;
    ctx.fillRect(x + 18, legY, 10, 60);
    ctx.fillRect(x + w - 28, legY, 10, 60);
    ctx.fillRect(x + 46, legY, 8, 56);
    ctx.fillRect(x + w - 54, legY, 8, 56);

    // small mug for fun (kopp)
    ctx.fillStyle = "#dbeafe";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 3;
    const mx = x + w - 54, my = topY + 10;
    roundRectStroke(mx, my, 22, 20, 7);
    ctx.fillRect(mx+3, my+3, 16, 14);
    ctx.beginPath();
    ctx.arc(mx + 22, my + 11, 7, -0.7, 0.7);
    ctx.stroke();
  }

  function drawChair(){
    const x = props.chair.x, w = props.chair.w;
    const seatY = FLOOR_Y - 52;

    // shadow
    ctx.save();
    ctx.fillStyle = "rgba(17,24,39,0.18)";
    ctx.beginPath();
    ctx.ellipse(x + w/2, FLOOR_Y + 2, w*0.34, 8, 0, 0, Math.PI*2);
    ctx.fill();
    ctx.restore();

    // cushion
    ctx.fillStyle = "#fef3c7";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 3;
    roundRectStroke(x + 10, seatY, w - 20, 18, 10);
    ctx.fillRect(x + 12, seatY + 2, w - 24, 14);

    // seat base
    ctx.fillStyle = "#d97706";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    roundRectStroke(x + 6, seatY + 16, w - 12, 18, 10);
    ctx.fillRect(x + 10, seatY + 20, w - 20, 10);

    // legs
    ctx.fillStyle = "#9ca3af";
    ctx.fillRect(x + 14, seatY + 34, 10, 36);
    ctx.fillRect(x + w - 24, seatY + 34, 10, 36);

    // backrest
    ctx.fillStyle = "#e5e7eb";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    roundRectStroke(x + 10, seatY - 74, w - 20, 70, 14);
    ctx.fillRect(x + 14, seatY - 70, w - 28, 62);

    // slats (gör den mer "stolig")
    ctx.save();
    ctx.globalAlpha = 0.45;
    ctx.fillStyle = "#111827";
    const slats = 4;
    for (let i=1;i<slats;i++){
      const sx = x + 14 + (i*(w - 28)/slats);
      ctx.fillRect(sx, seatY - 66, 2, 54);
    }
    ctx.restore();

    // label
    ctx.fillStyle = "rgba(17,24,39,0.85)";
    ctx.font = "1000 12px ui-monospace, monospace";
    ctx.fillText("STOL", x + w/2 - 18, seatY - 38);
  }

  function drawTree(tMs){
    const x = props.tree.x, w = props.tree.w;
    const topY = FLOOR_Y - 160;
    const pulse = 0.8 + 0.2 * Math.sin(tMs/250);

    // glowing aura
    ctx.save();
    ctx.globalAlpha = 0.12 * pulse;
    ctx.fillStyle = "#22c55e";
    ctx.beginPath();
    ctx.ellipse(x + w/2, FLOOR_Y - 85, 88, 110, 0, 0, Math.PI*2);
    ctx.fill();
    ctx.restore();

    // trunk
    ctx.fillStyle = "#78350f";
    ctx.fillRect(x + w/2 - 10, FLOOR_Y - 38, 20, 38);

    // tree layers
    const levels = 4;
    ctx.fillStyle = "#16a34a";
    ctx.strokeStyle = "#166534";
    ctx.lineWidth = 4;

    for(let i=0; i<levels; i++){
      const y1 = topY + i*32;
      const y2 = topY + (i+1)*32;
      const step = 16 + i*12;
      ctx.beginPath();
      ctx.moveTo(x + w/2, y1);
      ctx.lineTo(x + w/2 - step, y2);
      ctx.lineTo(x + w/2 + step, y2);
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    }

    // decorations
    const deco = [
      { x: x+28, y: topY+62, emoji: "🔴" },
      { x: x+72, y: topY+92, emoji: "🟠" },
      { x: x+48, y: topY+124, emoji: "🟡" }
    ];
    ctx.font = "24px " + getComputedStyle(document.body).fontFamily;
    for (const d of deco){
      ctx.fillText(d.emoji, d.x, d.y);
    }

    ctx.font = "28px " + getComputedStyle(document.body).fontFamily;
    ctx.fillText("⭐", x + w/2 - 14, topY + 10);
  }

  function drawOrnament(o){
    const y = floorYForRadius(o.r);
    ctx.beginPath();
    ctx.fillStyle = o.base;
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    ctx.arc(o.x, y, o.r, 0, Math.PI*2);
    ctx.fill();
    ctx.stroke();

    ctx.font = "26px " + getComputedStyle(document.body).fontFamily;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#111827";
    ctx.fillText(o.label, o.x, y + 1);
  }

  function drawJesper(tMs){
    const y = FLOOR_Y - 8;
    const moving = Math.abs(jesper.vx) > 18 && jesper.action !== "sit";
    const phase = (tMs/120) % (Math.PI*2);
    const walk = moving ? Math.sin(phase) : 0;
    const bob = moving ? Math.sin(phase*2)*2.0 : Math.sin(tMs/650)*1.3;

    const x = jesper.x;
    const face = jesper.facing || 1;

    const isKick = (jesper.action === "kick" && jesper.actionT < 0.28);
    const isBump = (jesper.action === "bump" && jesper.actionT < 0.40);
    const isWave = (jesper.action === "wave" && jesper.actionT < 1.20);
    const isSit  = (jesper.action === "sit");
    const isDance= (jesper.action === "dance" && jesper.actionT < 1.8);
    const isJuggle = (jesper.action === "juggle");
    const isDrop = (jesper.action === "drop");

    const sitDrop = isSit ? 18 : 0;

    ctx.save();
    ctx.translate(x, y + bob + sitDrop);
    ctx.scale(face, 1);

    // shake if bump
    if (isBump){
      ctx.translate((Math.random()-0.5)*5, (Math.random()-0.5)*3);
    }

    const hood = "#1f2937";
    const hoodie = "#111827";
    const skin = "#f8c7a1";
    const hair = "#111827";
    const hairH = "rgba(255,255,255,0.25)";

    // legs
    const legW = 9;
    const legH = isSit ? 8 : 18;
    const legA = isSit ? 0 : walk*6;
    const legB = isSit ? 0 : -walk*6;
    ctx.fillStyle = "#111827";
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;

    function leg(dx, a){
      ctx.save();
      ctx.translate(dx, 0);
      if (!isSit) ctx.rotate(a*0.02);
      ctx.beginPath();
      ctx.roundRect(-legW/2, -legH, legW, legH, 4);
      ctx.fill();
      ctx.restore();
    }
    leg(-10, legA);
    leg( 10, legB);

    // body
    ctx.fillStyle = hoodie;
    ctx.beginPath();
    ctx.roundRect(-24, -44, 48, 42, 18);
    ctx.fill();
    ctx.stroke();

    // belly highlight
    ctx.fillStyle = "rgba(255,255,255,0.08)";
    ctx.beginPath();
    ctx.ellipse(0, -24, 16, 12, 0, 0, Math.PI*2);
    ctx.fill();

    // arms
    const arm = moving ? Math.sin(phase*2)*5 : Math.sin(tMs/550)*2;
    const kickArm = isKick ? -14 : 0;
    const danceArm = isDance ? Math.sin(tMs/120)*10 : 0;

    function drawArm(x1,y1,x2,y2,color){
      ctx.save();
      ctx.strokeStyle = "#111827";
      ctx.lineWidth = 7;
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.moveTo(x1,y1);
      ctx.lineTo(x2,y2);
      ctx.stroke();
      ctx.restore();

      ctx.save();
      ctx.strokeStyle = color;
      ctx.lineWidth = 11;
      ctx.lineCap = "round";
      ctx.beginPath();
      ctx.moveTo(x1,y1);
      ctx.lineTo(x2,y2);
      ctx.stroke();
      ctx.restore();
    }

    // left arm
    if (isJuggle){
      drawArm(-18, -6, -34, -22 + Math.sin(tMs/170)*6, hood);
    } else if (isDance){
      drawArm(-18, -6, -34, -4 + danceArm, hood);
    } else {
      drawArm(-18, -6, -30,  6 - arm, hood);
    }

    // right arm
    if (isWave){
      drawArm(18, -6, 30, -16 + Math.sin(tMs/150)*10, hood);
    } else if (isJuggle){
      drawArm(18, -6, 34, -22 + Math.sin(tMs/160+1.2)*6, hood);
    } else if (isDance){
      drawArm(18, -6, 34, -4 - danceArm, hood);
    } else {
      drawArm( 18, -6,  30, 6 + arm + kickArm, hood);
    }

    // head (tilt)
    const tilt = clamp(jesper.tilt||0, -1, 1) * 0.12;
    ctx.save();
    ctx.translate(0, -34);
    ctx.rotate(tilt);
    ctx.translate(0, 34);

    ctx.beginPath();
    ctx.fillStyle = skin;
    ctx.strokeStyle = "#111827";
    ctx.lineWidth = 4;
    ctx.arc(0, -34, 16, 0, Math.PI*2);
    ctx.fill();
    ctx.stroke();

    // hair cap
    ctx.beginPath();
    ctx.fillStyle = hair;
    ctx.arc(0, -44, 17, Math.PI, 0);
    ctx.closePath();
    ctx.fill();

    ctx.strokeStyle = hairH;
    ctx.lineWidth = 2.2;
    for (let i=-12; i<=12; i+=6){
      ctx.beginPath();
      ctx.arc(i, -50, 3.6, 0, Math.PI*2);
      ctx.stroke();
    }

    // eyes
    const blink = (jesper.blinkT < 0.12);
    const excited = (jesper.faceMood === "star" || jesper.faceMood === "ecstatic");
    const browUp = (jesper.browT||0) > 0;

    function circleFill(cx, cy, r){
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, Math.PI*2);
      ctx.fill();
    }

    if (excited){
      // star eyes
      ctx.strokeStyle = "#111827";
      ctx.lineWidth = 2.4;
      for (const sx of [-6, 6]){
        ctx.beginPath();
        ctx.moveTo(sx-3, -37); ctx.lineTo(sx+3, -35);
        ctx.moveTo(sx+3, -37); ctx.lineTo(sx-3, -35);
        ctx.stroke();
      }
    } else if (blink){
      ctx.strokeStyle = "rgba(17,24,39,0.8)";
      ctx.lineWidth = 2.6;
      ctx.beginPath();
      ctx.moveTo(-8, -36); ctx.lineTo(-2, -36);
      ctx.moveTo( 2, -36); ctx.lineTo( 8, -36);
      ctx.stroke();
    } else {
      ctx.fillStyle = "#111827";
      circleFill(-5, -36, 2.2);
      circleFill( 5, -36, 2.2);
    }

    // eyebrows (micro reaction)
    ctx.strokeStyle = "rgba(17,24,39,0.75)";
    ctx.lineWidth = 2.3;
    const by = -41 - (browUp ? 2 : 0);
    ctx.beginPath();
    ctx.moveTo(-9, by); ctx.lineTo(-2, by-1);
    ctx.moveTo( 2, by-1); ctx.lineTo( 9, by);
    ctx.stroke();

    // mouth (neutral default)
    const mood = jesper.faceMood || "neutral";
    ctx.strokeStyle = "rgba(17,24,39,0.88)";
    ctx.lineWidth = 2.7;
    ctx.beginPath();
    if (mood === "neutral"){
      ctx.moveTo(-6, -26); ctx.lineTo(6, -26);
    } else if (mood === "smile"){
      ctx.arc(0, -27, 4, 0.15*Math.PI, 0.85*Math.PI);
    } else if (mood === "grin"){
      ctx.arc(0, -27, 5.4, 0.10*Math.PI, 0.90*Math.PI);
    } else if (mood === "smirk"){
      ctx.arc(1.2, -27, 4.4, 0.20*Math.PI, 0.95*Math.PI);
    } else if (mood === "frown"){
      ctx.arc(0, -23, 4.2, 1.15*Math.PI, 1.85*Math.PI);
    } else if (mood === "ecstatic"){
      ctx.arc(0, -28, 6.3, 0.05*Math.PI, 0.95*Math.PI);
    } else if (isSit){
      ctx.arc(0, -26, 4.0, 0.25*Math.PI, 0.75*Math.PI);
    } else {
      ctx.moveTo(-6, -26); ctx.lineTo(6, -26);
    }
    ctx.stroke();

    ctx.restore(); // head tilt

    // juggle balls
    if (isJuggle || isDrop){
      const t = tMs/240;
      for (let i=0;i<3;i++){
        const ph = t + i*(Math.PI*2/3);
        const bx = Math.sin(ph)*14;
        const by2 = -68 - (Math.cos(ph)*10);
        ctx.save();
        ctx.fillStyle = ["#ef4444","#22c55e","#3b82f6"][i];
        ctx.strokeStyle = "#111827";
        ctx.lineWidth = 2.4;
        ctx.beginPath();
        ctx.arc(bx, by2, 5.2, 0, Math.PI*2);
        ctx.fill();
        ctx.stroke();
        ctx.restore();
      }
    }

    // dance feet sparkle
    if (isDance){
      ctx.save();
      ctx.globalAlpha = 0.45;
      ctx.fillStyle = "#111827";
      for (let i=0;i<4;i++){
        ctx.fillRect(-18 + i*12, -4, 2, 2);
      }
      ctx.restore();
    }

    // label
    ctx.fillStyle = "rgba(17,24,39,0.55)";
    ctx.font = "900 12px ui-monospace, monospace";
    ctx.textAlign = "center";
    ctx.fillText("JESPER", 0, -62);

    ctx.restore();
  }

  function drawJoystick(){
    const bx = clamp(state.joy.startX, ROOM.x+40, ROOM.x+ROOM.w-40);
    const by = ROOM.y + ROOM.h - 80;
    const kx = bx + state.joy.dx;

    ctx.globalAlpha = 0.9;
    ctx.fillStyle = "rgba(17,24,39,0.07)";
    ctx.strokeStyle = "rgba(17,24,39,0.35)";
    ctx.lineWidth = 3;
    ctx.beginPath(); ctx.arc(bx, by, 24, 0, Math.PI*2); ctx.fill(); ctx.stroke();

    ctx.fillStyle = "rgba(37,99,235,0.18)";
    ctx.strokeStyle = "rgba(17,24,39,0.55)";
    ctx.beginPath(); ctx.arc(kx, by, 16, 0, Math.PI*2); ctx.fill(); ctx.stroke();
    ctx.globalAlpha = 1;
  }

  // ---------- Loop ----------
  let last = nowMs();
  function frame(t){
    const dt = Math.min(0.033, (t - last) / 1000);
    last = t;

    update(dt, t);
    draw(t);

    requestAnimationFrame(frame);
  }

  try{
    setTimeout(()=>toast("⏰ → 🍬 → ⭐ och sitt ner."), 900);
    requestAnimationFrame(frame);
  } catch(err){
    console.error(err);
    toast("Krasch 😵", 4000);
  }

})();


