(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  let game, scene, loading = false, muted = true, audio;
  let best = 0;
  try { best = Number(localStorage.getItem("kunhwa.orbit.best") || 0); } catch { /* Private storage is optional. */ }
  $("orbitBest").textContent = best;
  const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const W = innerWidth <= 800 ? 480 : 960, H = innerWidth <= 800 ? 640 : 600;
  function tone(frequency, duration = .09) {
    if (muted) return;
    try {
      audio ||= new (window.AudioContext || window.webkitAudioContext)(); audio.resume();
      const oscillator = audio.createOscillator(), gain = audio.createGain();
      oscillator.connect(gain); gain.connect(audio.destination); oscillator.frequency.value = frequency;
      gain.gain.setValueAtTime(.055, audio.currentTime); gain.gain.exponentialRampToValueAtTime(.001, audio.currentTime + duration);
      oscillator.start(); oscillator.stop(audio.currentTime + duration);
    } catch { muted = true; }
  }
  function icon(button, name) { button.innerHTML = window.toolsIcon(name); }
  function pause(value, automatic = false) {
    if (!scene?.running) return;
    scene.paused = value;
    if (value) scene.physics.pause(); else scene.physics.resume();
    icon($("orbitPause"), value ? "play-fill" : "pause-fill");
    $("orbitPause").setAttribute("aria-label", value ? "계속하기" : "일시 정지");
    $("orbitPause").title = value ? "계속하기" : "일시 정지";
    $("orbitStatus").textContent = value ? (automatic ? "비행 일시 정지" : "일시 정지") : "비행 중";
    $("orbitOverlay").hidden = !value;
    if (value) { $("orbitMessage").textContent = "잠시 정박 중"; $("orbitStart").textContent = "계속하기"; }
  }
  async function open() {
    if (loading || game) return;
    loading = true;
    try {
      if (!window.Phaser) await new Promise((resolve, reject) => {
        const script = document.createElement("script"); script.src = "./assets/vendor/phaser.min.js"; script.onload = resolve; script.onerror = reject; document.head.append(script);
      });
      class Flight extends Phaser.Scene {
        constructor() { super("flight"); this.running = false; this.paused = false; }
        preload() { this.load.svg("ship", "./assets/orbit-ship.svg", { width: 100, height: 90 }); this.load.svg("rock", "./assets/orbit-rock.svg", { width: 100, height: 100 }); this.load.svg("crystal", "./assets/orbit-crystal.svg", { width: 60, height: 80 }); }
        create() {
          scene = this;
          this.stars = Array.from({ length: 100 }, () => this.add.rectangle(Phaser.Math.Between(0, W), Phaser.Math.Between(0, H), Phaser.Math.Between(1, 2), 2, 0xe2f3ef, Phaser.Math.FloatBetween(.12, .65)));
          this.add.ellipse(W / 2, -85, W * .9, 450, 0x112c32).setStrokeStyle(1, 0x244a51);
          this.add.ellipse(W / 2, -160, W * .65, 440, 0x16232d).setStrokeStyle(1, 0x32404c);
          this.ship = this.physics.add.image(W / 2, H * .8, "ship").setDisplaySize(48, 44).setCollideWorldBounds(true);
          this.ship.body.setSize(44, 56).setOffset(28, 16);
          this.rocks = this.physics.add.group(); this.crystals = this.physics.add.group();
          this.physics.add.overlap(this.ship, this.rocks, (_, rock) => this.hit(rock));
          this.physics.add.overlap(this.ship, this.crystals, (_, crystal) => this.collect(crystal));
          this.keys = this.input.keyboard.createCursorKeys(); this.wasd = this.input.keyboard.addKeys("W,A,S,D");
          this.input.keyboard.on("keydown-SPACE", (event) => { if (location.hash === "#gameOrbit") { event.preventDefault(); if (this.running) pause(!this.paused); else this.startFlight(); } });
          this.input.on("pointerdown", (pointer) => { if (this.running && !this.paused) this.target = { x: pointer.x, y: pointer.y }; });
          this.input.on("pointermove", (pointer) => { if (pointer.isDown && this.running && !this.paused) this.target = { x: pointer.x, y: pointer.y }; });
          this.input.on("pointerup", () => { this.target = null; });
          this.input.keyboard.enabled = location.hash === "#gameOrbit";
          this.game.input.keyboard.enabled = location.hash === "#gameOrbit";
          $("orbitStart").disabled = false; $("orbitStatus").textContent = "비행 준비";
        }
        startFlight() {
          this.rocks.clear(true, true); this.crystals.clear(true, true); this.physics.resume();
          Object.assign(this, { running: true, paused: false, elapsed: 0, score: 0, lives: 3, cooldown: 0, rockClock: 0, crystalClock: 0, target: null });
          this.ship.setPosition(W / 2, H * .8).setAlpha(1).setVelocity(0); $("orbitOverlay").hidden = true; $("orbitPause").disabled = false;
          icon($("orbitPause"), "pause-fill"); $("orbitStatus").textContent = "비행 중"; this.hud(); tone(420);
        }
        hud() { $("orbitScore").textContent = Math.floor(this.score); $("orbitLives").textContent = this.lives; $("orbitLevel").textContent = `SECTOR ${String(1 + Math.floor(this.elapsed / 20000)).padStart(2, "0")}`; }
        collect(crystal) {
          crystal.destroy(); this.score += 100; this.hud(); tone(760);
          if (!reduced) { const ring = this.add.circle(this.ship.x, this.ship.y, 12).setStrokeStyle(2, 0x6effd3); this.tweens.add({ targets: ring, scale: 3, alpha: 0, duration: 250, onComplete: () => ring.destroy() }); }
        }
        hit(rock) {
          if (this.cooldown > 0 || !this.running) return;
          rock.destroy(); this.lives--; this.cooldown = 1500; tone(120, .2); if (!reduced) this.cameras.main.shake(120, .004);
          this.hud(); if (!this.lives) this.finish();
        }
        finish() {
          this.running = false; this.physics.pause(); this.ship.setAlpha(1);
          best = Math.max(best, Math.floor(this.score)); $("orbitBest").textContent = best;
          try { localStorage.setItem("kunhwa.orbit.best", String(best)); } catch { /* No persistence in restricted storage. */ }
          $("orbitOverlay").hidden = false; $("orbitMessage").textContent = `${Math.floor(this.score)}점 · 비행 종료`; $("orbitStart").textContent = "다시 비행"; $("orbitPause").disabled = true; $("orbitStatus").textContent = "비행 종료";
        }
        update(_, delta) {
          if (!this.running || this.paused || location.hash !== "#gameOrbit") return;
          const dt = Math.min(delta, 40); this.elapsed += dt; this.score += dt * .012; this.cooldown = Math.max(0, this.cooldown - dt);
          this.ship.setAlpha(this.cooldown > 0 ? (.6 + .4 * Math.sin(this.elapsed / 70)) : 1);
          const x = Number(this.keys.right.isDown || this.wasd.D.isDown) - Number(this.keys.left.isDown || this.wasd.A.isDown);
          const y = Number(this.keys.down.isDown || this.wasd.S.isDown) - Number(this.keys.up.isDown || this.wasd.W.isDown);
          if (x || y) { this.target = null; this.ship.setVelocity(x * 340, y * 340); }
          else if (this.target) this.physics.moveTo(this.ship, this.target.x, this.target.y, 480);
          else this.ship.setVelocity(0);
          if (this.target && Phaser.Math.Distance.Between(this.ship.x, this.ship.y, this.target.x, this.target.y) < 8) this.ship.setVelocity(0);
          this.ship.setAngle(x * 12);
          this.stars.forEach((star) => { star.y += dt * .05; if (star.y > H) { star.y = 0; star.x = Phaser.Math.Between(0, W); } });
          this.rockClock += dt; this.crystalClock += dt;
          const speed = Math.min(380, 125 + this.elapsed / 500), interval = Math.max(330, 950 - this.elapsed / 120);
          if (this.rockClock > interval) {
            this.rockClock = 0;
            const rock = this.rocks.create(Phaser.Math.Between(40, W - 40), -60, "rock"), size = Phaser.Math.Between(36, 80);
            rock.setDisplaySize(size, size).setVelocity(Phaser.Math.Between(-30, 30), speed).setAngularVelocity(Phaser.Math.Between(-50, 50)); rock.body.setCircle(39, 11, 11);
          }
          if (this.crystalClock > 1600) { this.crystalClock = 0; this.crystals.create(Phaser.Math.Between(50, W - 50), -30, "crystal").setDisplaySize(22, 30).setVelocity(0, speed * .9); }
          [this.rocks, this.crystals].forEach((group) => group.getChildren().forEach((item) => { if (item.y > H + 80) item.destroy(); }));
          this.hud();
        }
      }
      game = new Phaser.Game({ type: Phaser.AUTO, parent: "orbitCanvas", width: W, height: H, backgroundColor: "#0b1019", scene: Flight, physics: { default: "arcade", arcade: { debug: false } }, scale: { mode: Phaser.Scale.FIT, autoCenter: Phaser.Scale.CENTER_BOTH }, audio: { noAudio: true }, input: { keyboard: true }, render: { antialias: true } });
    } catch { $("orbitStatus").textContent = "게임 파일을 불러오지 못했습니다. 페이지를 다시 열어 주세요."; }
    finally { loading = false; }
  }
  const enter = () => { location.hash = "gameOrbit"; open(); };
  $("brandMascot").addEventListener("dblclick", enter);
  $("brandMascot").addEventListener("keydown", (event) => { if (event.key === "Enter" && event.shiftKey) enter(); });
  $("orbitStart").onclick = () => { if (!scene) return; if (scene.running && scene.paused) pause(false); else scene.startFlight(); };
  $("orbitPause").onclick = () => pause(!scene?.paused);
  $("orbitSound").onclick = () => { muted = !muted; icon($("orbitSound"), muted ? "volume-mute-fill" : "volume-up-fill"); $("orbitSound").setAttribute("aria-label", muted ? "소리 켜기" : "소리 끄기"); $("orbitSound").title = muted ? "소리 켜기" : "소리 끄기"; if (!muted) tone(500); };
  const routeChanged = () => { const active = location.hash === "#gameOrbit"; if (active) open(); else pause(true, true); if (scene) { scene.input.keyboard.enabled = active; scene.game.input.keyboard.enabled = active; } };
  window.addEventListener("hashchange", routeChanged);
  window.addEventListener("tools:routechange", routeChanged);
  document.addEventListener("visibilitychange", () => { if (document.hidden) pause(true, true); });
  $("orbitStart").disabled = true;
  if (location.hash === "#gameOrbit") open();
})();
