(() => {
  "use strict";

  const canvas = document.getElementById("tableCanvas");
  const ctx = canvas.getContext("2d");
  const W = canvas.width;
  const H = canvas.height;
  const rail = 58;
  const ballRadius = 17;
  const bounds = { left: rail + ballRadius, right: W - rail - ballRadius, top: rail + ballRadius, bottom: H - rail - ballRadius };
  const initial = [
    { id: "cue", x: 300, y: 325, color: "#fffdf1", stroke: "#b7b4a8", label: "●" },
    { id: "yellow", x: 780, y: 230, color: "#f1d34f", stroke: "#987d16", label: "1" },
    { id: "red", x: 875, y: 410, color: "#dc554d", stroke: "#7c2725", label: "2" },
  ];

  let balls = cloneBalls(initial);
  let angle = 0;
  let power = 55;
  let sideSpin = 0;
  let verticalSpin = 0;
  let previewCushions = 3;
  let draggingBall = null;
  let running = false;
  let lastTime = 0;
  let shot = null;

  const angleInput = document.getElementById("angleInput");
  const powerInput = document.getElementById("powerInput");
  const cushionInput = document.getElementById("cushionInput");
  const spinPad = document.getElementById("spinPad");
  const spinMarker = document.getElementById("spinMarker");
  const stats = loadJSON("carom-lab-stats", { attempts: 0, successes: 0 });

  function cloneBalls(source) {
    return source.map((b) => ({ ...b, vx: 0, vy: 0, hit: false }));
  }

  function loadJSON(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; }
    catch { return fallback; }
  }

  function saveJSON(key, value) {
    localStorage.setItem(key, JSON.stringify(value));
  }

  function setStatus(kind, title, message) {
    const badge = document.getElementById("statusBadge");
    badge.className = `status-badge${kind ? ` ${kind}` : ""}`;
    badge.textContent = title;
    document.getElementById("statusText").textContent = message;
  }

  function updateControls() {
    angleInput.value = Math.round(angle);
    powerInput.value = power;
    cushionInput.value = String(previewCushions);
    document.getElementById("angleValue").textContent = `${Math.round(angle)}°`;
    document.getElementById("powerValue").textContent = `${power}%`;
    document.getElementById("sideSpinValue").textContent = sideSpin.toFixed(1);
    document.getElementById("verticalSpinValue").textContent = verticalSpin.toFixed(1);
    spinMarker.style.left = `${50 + sideSpin * 34}%`;
    spinMarker.style.top = `${50 - verticalSpin * 34}%`;
  }

  function updateStats() {
    document.getElementById("attemptCount").textContent = stats.attempts;
    document.getElementById("successCount").textContent = stats.successes;
    document.getElementById("successRate").textContent = stats.attempts ? `${Math.round(stats.successes / stats.attempts * 100)}%` : "0%";
    saveJSON("carom-lab-stats", stats);
  }

  function drawTable() {
    const g = ctx.createLinearGradient(0, 0, 0, H);
    g.addColorStop(0, "#a06b37");
    g.addColorStop(1, "#6d4224");
    roundRect(8, 8, W - 16, H - 16, 42, g);
    roundRect(rail - 12, rail - 12, W - (rail - 12) * 2, H - (rail - 12) * 2, 22, "#0c3f31");
    const cloth = ctx.createRadialGradient(W * .45, H * .35, 10, W * .5, H * .5, W * .65);
    cloth.addColorStop(0, "#20785d");
    cloth.addColorStop(1, "#0f553f");
    roundRect(rail, rail, W - rail * 2, H - rail * 2, 14, cloth);

    ctx.fillStyle = "#d9c28e";
    for (let i = 1; i < 8; i++) {
      marker(rail + (W - rail * 2) * i / 8, rail - 24);
      marker(rail + (W - rail * 2) * i / 8, H - rail + 24);
    }
    for (let i = 1; i < 4; i++) {
      marker(rail - 24, rail + (H - rail * 2) * i / 4);
      marker(W - rail + 24, rail + (H - rail * 2) * i / 4);
    }
  }

  function roundRect(x, y, w, h, r, fill) {
    ctx.beginPath();
    ctx.roundRect(x, y, w, h, r);
    ctx.fillStyle = fill;
    ctx.fill();
  }

  function marker(x, y) {
    ctx.save();
    ctx.translate(x, y);
    ctx.rotate(Math.PI / 4);
    ctx.fillRect(-4, -4, 8, 8);
    ctx.restore();
  }

  function drawBall(ball) {
    ctx.save();
    ctx.shadowColor = "#001d1688";
    ctx.shadowBlur = 10;
    ctx.shadowOffsetY = 5;
    ctx.beginPath();
    ctx.arc(ball.x, ball.y, ballRadius, 0, Math.PI * 2);
    ctx.fillStyle = ball.color;
    ctx.fill();
    ctx.shadowColor = "transparent";
    ctx.lineWidth = 2;
    ctx.strokeStyle = ball.stroke;
    ctx.stroke();
    const shine = ctx.createRadialGradient(ball.x - 6, ball.y - 7, 1, ball.x - 3, ball.y - 4, 13);
    shine.addColorStop(0, "#ffffffdd");
    shine.addColorStop(1, "#ffffff00");
    ctx.fillStyle = shine;
    ctx.fill();
    if (ball.id !== "cue") {
      ctx.fillStyle = ball.id === "red" ? "#fff" : "#2b260f";
      ctx.font = "700 13px Segoe UI";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(ball.label, ball.x, ball.y + .5);
    } else {
      const radians = angle * Math.PI / 180;
      ctx.beginPath();
      ctx.arc(ball.x + Math.cos(radians + Math.PI / 2) * sideSpin * 7, ball.y + Math.sin(radians + Math.PI / 2) * sideSpin * 7 - verticalSpin * 7, 2.6, 0, Math.PI * 2);
      ctx.fillStyle = "#d63e36";
      ctx.fill();
    }
    ctx.restore();
  }

  function drawPrediction() {
    if (running) return;
    const cue = balls[0];
    const radians = angle * Math.PI / 180;
    let origin = { x: cue.x, y: cue.y };
    let direction = { x: Math.cos(radians), y: Math.sin(radians) };
    const segments = [];
    let hitBall = null;

    for (let bounce = 0; bounce <= previewCushions; bounce++) {
      const railHit = nextRailHit(origin, direction);
      const objectHit = nextObjectHit(origin, direction, balls.slice(1));
      if (objectHit && objectHit.t < railHit.t) {
        segments.push({ from: origin, to: objectHit.point, type: "ball" });
        hitBall = objectHit.ball;
        drawObjectExit(hitBall, direction, objectHit.point);
        break;
      }
      segments.push({ from: origin, to: railHit.point, type: "rail" });
      if (bounce === previewCushions) break;
      origin = { x: railHit.point.x + railHit.reflected.x * .5, y: railHit.point.y + railHit.reflected.y * .5 };
      direction = applyPreviewSpin(railHit.reflected, railHit.axis);
    }

    ctx.save();
    ctx.lineWidth = 3;
    ctx.strokeStyle = "#d8f57c";
    ctx.setLineDash([10, 8]);
    ctx.beginPath();
    ctx.moveTo(segments[0].from.x, segments[0].from.y);
    segments.forEach((segment) => ctx.lineTo(segment.to.x, segment.to.y));
    ctx.stroke();
    ctx.setLineDash([]);
    segments.filter((s) => s.type === "rail").forEach((s, index) => {
      ctx.beginPath();
      ctx.arc(s.to.x, s.to.y, 10, 0, Math.PI * 2);
      ctx.fillStyle = "#c8e66b";
      ctx.fill();
      ctx.fillStyle = "#173425";
      ctx.font = "700 11px Segoe UI";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(String(index + 1), s.to.x, s.to.y + .5);
    });
    if (hitBall) {
      ctx.beginPath();
      ctx.arc(hitBall.x, hitBall.y, ballRadius + 7, 0, Math.PI * 2);
      ctx.strokeStyle = "#fff4";
      ctx.lineWidth = 2;
      ctx.stroke();
    }
    ctx.restore();
  }

  function nextRailHit(o, d) {
    const candidates = [];
    if (d.x > 0) candidates.push({ t: (bounds.right - o.x) / d.x, axis: "x" });
    if (d.x < 0) candidates.push({ t: (bounds.left - o.x) / d.x, axis: "x" });
    if (d.y > 0) candidates.push({ t: (bounds.bottom - o.y) / d.y, axis: "y" });
    if (d.y < 0) candidates.push({ t: (bounds.top - o.y) / d.y, axis: "y" });
    const hit = candidates.filter((c) => c.t > .1).sort((a, b) => a.t - b.t)[0];
    const reflected = { x: hit.axis === "x" ? -d.x : d.x, y: hit.axis === "y" ? -d.y : d.y };
    return { ...hit, point: { x: o.x + d.x * hit.t, y: o.y + d.y * hit.t }, reflected };
  }

  function nextObjectHit(o, d, objects) {
    let closest = null;
    objects.forEach((ball) => {
      const ox = o.x - ball.x;
      const oy = o.y - ball.y;
      const b = 2 * (ox * d.x + oy * d.y);
      const c = ox * ox + oy * oy - (ballRadius * 2) ** 2;
      const disc = b * b - 4 * c;
      if (disc < 0) return;
      const t = (-b - Math.sqrt(disc)) / 2;
      if (t > 1 && (!closest || t < closest.t)) closest = { t, ball, point: { x: o.x + d.x * t, y: o.y + d.y * t } };
    });
    return closest;
  }

  function drawObjectExit(ball, incoming, contact) {
    const nx = ball.x - contact.x;
    const ny = ball.y - contact.y;
    const length = Math.hypot(nx, ny) || 1;
    ctx.save();
    ctx.strokeStyle = "#ffffff88";
    ctx.lineWidth = 2;
    ctx.setLineDash([5, 6]);
    ctx.beginPath();
    ctx.moveTo(ball.x, ball.y);
    ctx.lineTo(ball.x + nx / length * 90, ball.y + ny / length * 90);
    ctx.stroke();
    ctx.restore();
  }

  function applyPreviewSpin(direction, axis) {
    const spinAngle = sideSpin * .055 * (axis === "x" ? Math.sign(direction.x || 1) : -Math.sign(direction.y || 1));
    const cos = Math.cos(spinAngle);
    const sin = Math.sin(spinAngle);
    return { x: direction.x * cos - direction.y * sin, y: direction.x * sin + direction.y * cos };
  }

  function render() {
    ctx.clearRect(0, 0, W, H);
    drawTable();
    drawPrediction();
    balls.forEach(drawBall);
  }

  function pointerPosition(event) {
    const rect = canvas.getBoundingClientRect();
    return { x: (event.clientX - rect.left) * W / rect.width, y: (event.clientY - rect.top) * H / rect.height };
  }

  function ballAt(point) {
    return balls.find((ball) => Math.hypot(ball.x - point.x, ball.y - point.y) <= ballRadius + 10);
  }

  function validPosition(target, x, y) {
    return !balls.some((b) => b !== target && Math.hypot(b.x - x, b.y - y) < ballRadius * 2 + 4);
  }

  canvas.addEventListener("pointerdown", (event) => {
    if (running) return;
    const point = pointerPosition(event);
    const selected = ballAt(point);
    if (selected) {
      draggingBall = selected;
      canvas.setPointerCapture(event.pointerId);
    } else {
      const cue = balls[0];
      angle = Math.atan2(point.y - cue.y, point.x - cue.x) * 180 / Math.PI;
      updateControls();
      render();
    }
  });

  canvas.addEventListener("pointermove", (event) => {
    if (!draggingBall || running) return;
    const point = pointerPosition(event);
    const x = Math.max(bounds.left, Math.min(bounds.right, point.x));
    const y = Math.max(bounds.top, Math.min(bounds.bottom, point.y));
    if (validPosition(draggingBall, x, y)) {
      draggingBall.x = x;
      draggingBall.y = y;
      render();
    }
  });

  canvas.addEventListener("pointerup", () => { draggingBall = null; });
  canvas.addEventListener("pointercancel", () => { draggingBall = null; });

  angleInput.addEventListener("input", () => { angle = Number(angleInput.value); updateControls(); render(); });
  powerInput.addEventListener("input", () => { power = Number(powerInput.value); updateControls(); });
  cushionInput.addEventListener("change", () => { previewCushions = Number(cushionInput.value); render(); });

  function setSpin(event) {
    const rect = spinPad.getBoundingClientRect();
    const x = ((event.clientX - rect.left) / rect.width - .5) * 2;
    const y = -((event.clientY - rect.top) / rect.height - .5) * 2;
    const length = Math.hypot(x, y);
    const scale = length > 1 ? 1 / length : 1;
    sideSpin = Math.round(x * scale * 10) / 10;
    verticalSpin = Math.round(y * scale * 10) / 10;
    updateControls();
    render();
  }
  spinPad.addEventListener("pointerdown", (event) => { spinPad.setPointerCapture(event.pointerId); setSpin(event); });
  spinPad.addEventListener("pointermove", (event) => { if (spinPad.hasPointerCapture(event.pointerId)) setSpin(event); });

  function shoot() {
    if (running) return;
    const speed = 290 + power * 5.1;
    const radians = angle * Math.PI / 180;
    balls.forEach((b) => { b.vx = 0; b.vy = 0; b.hit = false; });
    balls[0].vx = Math.cos(radians) * speed;
    balls[0].vy = Math.sin(radians) * speed;
    shot = { cushionCount: 0, touched: new Set(), finished: false };
    running = true;
    lastTime = performance.now();
    stats.attempts++;
    updateStats();
    setStatus("", "진행 중", "공의 실제 경로를 계산하고 있습니다.");
    requestAnimationFrame(step);
  }

  function step(now) {
    if (!running) return;
    const dt = Math.min((now - lastTime) / 1000, .025);
    lastTime = now;
    const substeps = 3;
    for (let s = 0; s < substeps; s++) simulate(dt / substeps);
    render();
    const moving = balls.some((b) => Math.hypot(b.vx, b.vy) > 5);
    if (moving) requestAnimationFrame(step);
    else finishShot();
  }

  function simulate(dt) {
    balls.forEach((ball) => {
      ball.x += ball.vx * dt;
      ball.y += ball.vy * dt;
      const speed = Math.hypot(ball.vx, ball.vy);
      const decel = 74 * (1 + Math.max(0, -verticalSpin) * .18);
      const nextSpeed = Math.max(0, speed - decel * dt);
      if (speed) { ball.vx *= nextSpeed / speed; ball.vy *= nextSpeed / speed; }
      let bounced = false;
      if (ball.x < bounds.left) { ball.x = bounds.left; ball.vx = Math.abs(ball.vx) * .88; ball.vy += sideSpin * 22; bounced = true; }
      if (ball.x > bounds.right) { ball.x = bounds.right; ball.vx = -Math.abs(ball.vx) * .88; ball.vy -= sideSpin * 22; bounced = true; }
      if (ball.y < bounds.top) { ball.y = bounds.top; ball.vy = Math.abs(ball.vy) * .88; ball.vx -= sideSpin * 22; bounced = true; }
      if (ball.y > bounds.bottom) { ball.y = bounds.bottom; ball.vy = -Math.abs(ball.vy) * .88; ball.vx += sideSpin * 22; bounced = true; }
      if (bounced && ball.id === "cue") shot.cushionCount++;
    });
    for (let i = 0; i < balls.length; i++) for (let j = i + 1; j < balls.length; j++) collide(balls[i], balls[j]);
  }

  function collide(a, b) {
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const distance = Math.hypot(dx, dy);
    if (!distance || distance >= ballRadius * 2) return;
    const nx = dx / distance;
    const ny = dy / distance;
    const overlap = ballRadius * 2 - distance;
    a.x -= nx * overlap / 2; a.y -= ny * overlap / 2;
    b.x += nx * overlap / 2; b.y += ny * overlap / 2;
    const relative = (a.vx - b.vx) * nx + (a.vy - b.vy) * ny;
    if (relative <= 0) return;
    const impulse = relative * .96;
    a.vx -= impulse * nx; a.vy -= impulse * ny;
    b.vx += impulse * nx; b.vy += impulse * ny;
    if (a.id === "cue" && b.id !== "cue") shot.touched.add(b.id);
    if (b.id === "cue" && a.id !== "cue") shot.touched.add(a.id);
  }

  function finishShot() {
    running = false;
    balls.forEach((b) => { b.vx = 0; b.vy = 0; });
    const success = shot.touched.size === 2 && shot.cushionCount >= 3;
    if (success) {
      stats.successes++;
      setStatus("success", "득점", `두 적구 접촉 · 쿠션 ${shot.cushionCount}회`);
    } else {
      const hitText = shot.touched.size === 0 ? "적구 미접촉" : shot.touched.size === 1 ? "적구 1개 접촉" : "두 적구 접촉";
      setStatus("fail", "실패", `${hitText} · 쿠션 ${shot.cushionCount}회`);
    }
    updateStats();
    render();
  }

  function randomize() {
    if (running) return;
    balls.forEach((ball, index) => {
      let x, y, tries = 0;
      do {
        x = bounds.left + 60 + Math.random() * (bounds.right - bounds.left - 120);
        y = bounds.top + 45 + Math.random() * (bounds.bottom - bounds.top - 90);
        tries++;
      } while (!validPosition(ball, x, y) && tries < 100);
      ball.x = x; ball.y = y;
    });
    setStatus("", "설계 중", "새로운 배치입니다. 예상 경로를 설계하세요.");
    render();
  }

  function reset() {
    if (running) return;
    balls = cloneBalls(initial);
    angle = 0; power = 55; sideSpin = 0; verticalSpin = 0; previewCushions = 3;
    setStatus("", "설계 중", "공을 드래그하거나 빈 곳을 눌러 조준하세요.");
    updateControls();
    render();
  }

  function savedLayouts() { return loadJSON("carom-lab-layouts", []); }
  function refreshLayouts(selected = "") {
    const select = document.getElementById("savedLayouts");
    const layouts = savedLayouts();
    select.innerHTML = '<option value="">저장된 배치 선택</option>';
    layouts.forEach((layout) => {
      const option = document.createElement("option");
      option.value = layout.id;
      option.textContent = layout.name;
      select.append(option);
    });
    select.value = selected;
  }

  document.getElementById("saveBtn").addEventListener("click", () => {
    const input = document.getElementById("layoutName");
    const name = input.value.trim() || `배치 ${new Date().toLocaleString("ko-KR", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })}`;
    const layouts = savedLayouts();
    const layout = { id: String(Date.now()), name, balls: balls.map(({ x, y }) => ({ x, y })), angle, power, sideSpin, verticalSpin, previewCushions };
    layouts.unshift(layout);
    saveJSON("carom-lab-layouts", layouts.slice(0, 30));
    input.value = "";
    refreshLayouts(layout.id);
    setStatus("", "저장 완료", `“${name}” 배치를 저장했습니다.`);
  });

  document.getElementById("loadBtn").addEventListener("click", () => {
    if (running) return;
    const id = document.getElementById("savedLayouts").value;
    const layout = savedLayouts().find((item) => item.id === id);
    if (!layout) return;
    balls.forEach((ball, i) => Object.assign(ball, layout.balls[i], { vx: 0, vy: 0 }));
    ({ angle, power, sideSpin, verticalSpin, previewCushions } = layout);
    updateControls(); render();
    setStatus("", "불러옴", `“${layout.name}” 배치를 불러왔습니다.`);
  });

  document.getElementById("deleteBtn").addEventListener("click", () => {
    const select = document.getElementById("savedLayouts");
    if (!select.value) return;
    saveJSON("carom-lab-layouts", savedLayouts().filter((item) => item.id !== select.value));
    refreshLayouts();
  });

  document.getElementById("shootBtn").addEventListener("click", shoot);
  document.getElementById("randomBtn").addEventListener("click", randomize);
  document.getElementById("resetBtn").addEventListener("click", reset);
  document.getElementById("clearStatsBtn").addEventListener("click", () => { stats.attempts = 0; stats.successes = 0; updateStats(); });
  window.addEventListener("resize", render);

  updateControls();
  updateStats();
  refreshLayouts();
  render();
})();
