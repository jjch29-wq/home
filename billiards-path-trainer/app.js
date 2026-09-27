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
  let physicsAccumulator = 0;
  let shot = null;
  let firstBallId = "yellow";
  let successPath = null;
  let solving = false;
  let demoPosition = null;
  let solveRunId = 0;
  let lastShotSetup = null;

  const angleInput = document.getElementById("angleInput");
  const powerInput = document.getElementById("powerInput");
  const cushionInput = document.getElementById("cushionInput");
  const spinPad = document.getElementById("spinPad");
  const spinMarker = document.getElementById("spinMarker");
  const stats = loadJSON("carom-lab-stats", { attempts: 0, successes: 0 });

  function cloneBalls(source) {
    return source.map((b) => ({ ...b, vx: 0, vy: 0, hit: false }));
  }

  function captureShotSetup(path = null) {
    return {
      balls: cloneBalls(balls),
      successPath: path,
      angle,
      power,
      sideSpin,
      verticalSpin,
      previewCushions,
      firstBallId,
    };
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

  function clearSuccessPath(preserveRetry = false) {
    successPath = null;
    demoPosition = null;
    document.getElementById("successShootBtn").disabled = true;
    updateFirstHitGuide(null);
    if (!preserveRetry) {
      lastShotSetup = null;
      document.getElementById("retryShotBtn").disabled = true;
    }
  }

  function updateFirstHitGuide(path) {
    const event = path?.events?.find((item) => item.kind === "ball");
    const ball = document.getElementById("firstHitBall");
    const marker = document.getElementById("firstHitMarker");
    const thicknessText = document.getElementById("firstHitThickness");
    const directionText = document.getElementById("firstHitDirection");
    ball.classList.toggle("yellow", firstBallId === "yellow");
    ball.classList.toggle("red", firstBallId === "red");
    if (!event) {
      marker.hidden = true;
      thicknessText.textContent = "경로를 먼저 찾아주세요";
      directionText.textContent = "점이 표시된 곳을 맞히세요";
      return;
    }
    const thickness = event.thickness ?? 1;
    const eighths = Math.max(0, Math.min(8, Math.round(thickness * 8)));
    marker.hidden = false;
    marker.style.left = `${50 + (event.contactX ?? 0) * 38}%`;
    marker.style.top = `${50 + (event.contactY ?? 0) * 38}%`;
    thicknessText.textContent = `${event.label}공 ${eighths}/8 두께 (${Math.round(thickness * 100)}%)`;
    directionText.textContent = "파란 점 위치를 수구로 맞히세요";
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
    if (successPath) {
      drawSuccessPath();
      return;
    }
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

  function drawSuccessPath() {
    if (!successPath || successPath.points.length < 2) return;
    ctx.save();
    ctx.strokeStyle = "#7de8ff";
    ctx.lineWidth = 4;
    ctx.setLineDash([12, 7]);
    ctx.beginPath();
    ctx.moveTo(successPath.points[0].x, successPath.points[0].y);
    successPath.points.slice(1).forEach((point) => ctx.lineTo(point.x, point.y));
    ctx.stroke();
    ctx.setLineDash([]);
    successPath.events.forEach((event) => {
      if (event.kind === "ball") {
        const labelOnLeft = event.x > bounds.right - 90;
        const labelX = event.x + (labelOnLeft ? -60 : 60);
        const labelY = Math.max(bounds.top + 12, Math.min(bounds.bottom - 12, event.y + 25));

        // 목적구를 맞히는 순간의 수구 위치와 두께를 보여 주는 고스트볼.
        ctx.beginPath();
        ctx.arc(event.x, event.y, ballRadius, 0, Math.PI * 2);
        ctx.fillStyle = "#ffffff4d";
        ctx.fill();
        ctx.strokeStyle = "#ffffffcc";
        ctx.lineWidth = 2;
        ctx.setLineDash([4, 3]);
        ctx.stroke();
        ctx.setLineDash([]);

        ctx.beginPath();
        ctx.roundRect(labelX - 44, labelY - 10, 88, 20, 10);
        ctx.fillStyle = "#082f27dd";
        ctx.fill();
        ctx.strokeStyle = "#7de8ffaa";
        ctx.lineWidth = 1;
        ctx.stroke();
        ctx.fillStyle = "#ffffff";
        ctx.font = "700 10px Segoe UI";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        const thickness = event.thickness ?? 1;
        const eighths = Math.max(0, Math.min(8, Math.round(thickness * 8)));
        ctx.fillText(`${event.label} 두께 ${eighths}/8 (${Math.round(thickness * 100)}%)`, labelX, labelY + .5);
      }
      ctx.beginPath();
      ctx.arc(event.x, event.y, 12, 0, Math.PI * 2);
      ctx.fillStyle = event.kind === "ball" ? "#ffffff" : "#7de8ff";
      ctx.fill();
      ctx.fillStyle = "#12372f";
      ctx.font = "700 11px Segoe UI";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(event.label, event.x, event.y + .5);
    });
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
    balls.forEach((ball) => {
      if (ball.id === "cue" && demoPosition) drawBall({ ...ball, x: demoPosition.x, y: demoPosition.y });
      else drawBall(ball);
    });
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
    if (running || solving) return;
    const point = pointerPosition(event);
    const selected = ballAt(point);
    if (selected) {
      draggingBall = selected;
      clearSuccessPath();
      canvas.setPointerCapture(event.pointerId);
    } else {
      const cue = balls[0];
      angle = Math.atan2(point.y - cue.y, point.x - cue.x) * 180 / Math.PI;
      clearSuccessPath();
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

  angleInput.addEventListener("input", () => { angle = Number(angleInput.value); clearSuccessPath(); updateControls(); render(); });
  powerInput.addEventListener("input", () => { power = Number(powerInput.value); clearSuccessPath(); updateControls(); });
  cushionInput.addEventListener("change", () => { previewCushions = Number(cushionInput.value); clearSuccessPath(); render(); });

  function setSpin(event) {
    const rect = spinPad.getBoundingClientRect();
    const x = ((event.clientX - rect.left) / rect.width - .5) * 2;
    const y = -((event.clientY - rect.top) / rect.height - .5) * 2;
    const length = Math.hypot(x, y);
    const scale = length > 1 ? 1 / length : 1;
    sideSpin = Math.round(x * scale * 10) / 10;
    verticalSpin = Math.round(y * scale * 10) / 10;
    clearSuccessPath();
    updateControls();
    render();
  }
  spinPad.addEventListener("pointerdown", (event) => { spinPad.setPointerCapture(event.pointerId); setSpin(event); });
  spinPad.addEventListener("pointermove", (event) => { if (spinPad.hasPointerCapture(event.pointerId)) setSpin(event); });

  function shoot() {
    if (running || solving) return;
    if (successPath?.frames?.length) {
      playSolvedPhysicsShot();
      return;
    }
    clearSuccessPath();
    lastShotSetup = captureShotSetup();
    const speed = 290 + power * 5.1;
    const radians = angle * Math.PI / 180;
    balls.forEach((b) => { b.vx = 0; b.vy = 0; b.hit = false; });
    balls[0].vx = Math.cos(radians) * speed;
    balls[0].vy = Math.sin(radians) * speed;
    shot = { cushionCount: 0, touched: new Set(), invalidOrder: false, secondTouchedAtCushions: null, scored: false };
    running = true;
    lastTime = performance.now();
    physicsAccumulator = 0;
    stats.attempts++;
    updateStats();
    setStatus("", "진행 중", "공의 실제 경로를 계산하고 있습니다.");
    requestAnimationFrame(step);
  }

  function playSolvedPhysicsShot() {
    if (!successPath?.frames?.length || running || solving) return;
    const solved = successPath;
    lastShotSetup = captureShotSetup(solved);
    document.getElementById("retryShotBtn").disabled = true;
    const frames = solved.frames;
    const frameStep = Math.max(1, Math.ceil(frames.length / 420));
    let frameIndex = 0;
    running = true;
    stats.attempts++;
    updateStats();
    setStatus("", "실제 샷 실행", "고정시간 물리 계산 재생 0%");

    const playbackTimer = window.setInterval(() => {
      const frame = frames[Math.min(frameIndex, frames.length - 1)];
      balls.forEach((ball, ballIndex) => {
        ball.x = frame[ballIndex].x;
        ball.y = frame[ballIndex].y;
        ball.vx = frame[ballIndex].vx;
        ball.vy = frame[ballIndex].vy;
      });
      render();
      frameIndex += frameStep;
      const progress = Math.min(100, Math.round(frameIndex / frames.length * 100));
      document.getElementById("statusText").textContent = `고정시간 물리 계산 재생 ${progress}%`;
      if (frameIndex < frames.length) {
        return;
      }
      window.clearInterval(playbackTimer);
      running = false;
      balls.forEach((ball) => { ball.vx = 0; ball.vy = 0; });
      stats.successes++;
      updateStats();
      setStatus("success", "실제 샷 성공", `${firstBallId === "yellow" ? "노란공" : "빨간공"} 먼저 · 두 번째 적구 전 ${solved.cushions}쿠션`);
      clearSuccessPath(true);
      document.getElementById("retryShotBtn").disabled = false;
      render();
    }, 1000 / 60);
  }

  function playSuccessRoute() {
    if (!successPath || running || solving) return;
    const route = successPath;
    const segments = [];
    let totalLength = 0;
    for (let i = 1; i < route.points.length; i++) {
      const from = route.points[i - 1];
      const to = route.points[i];
      const length = Math.hypot(to.x - from.x, to.y - from.y);
      if (length > .1) {
        segments.push({ from, to, length, start: totalLength });
        totalLength += length;
      }
    }
    if (!segments.length) return;

    running = true;
    stats.attempts++;
    updateStats();
    const duration = Math.max(1800, Math.min(6500, totalLength / (360 + route.power * 2.4) * 1000));
    const startedAt = performance.now();
    setStatus("", "경로 재생", `${firstBallId === "yellow" ? "노란공" : "빨간공"}을 먼저 맞히는 ${route.cushions}쿠션 성공 경로입니다.`);

    function animateDemo(now) {
      const progress = Math.min(1, (now - startedAt) / duration);
      const traveled = totalLength * (1 - (1 - progress) ** 1.35);
      const segment = segments.find((item) => traveled <= item.start + item.length) || segments[segments.length - 1];
      const local = Math.max(0, Math.min(1, (traveled - segment.start) / segment.length));
      demoPosition = {
        x: segment.from.x + (segment.to.x - segment.from.x) * local,
        y: segment.from.y + (segment.to.y - segment.from.y) * local,
      };
      render();
      if (progress < 1) {
        requestAnimationFrame(animateDemo);
        return;
      }
      running = false;
      demoPosition = null;
      stats.successes++;
      updateStats();
      setStatus("success", "경로 재생 완료", `${firstBallId === "yellow" ? "노란공" : "빨간공"} 먼저 · ${route.cushions}쿠션 성공 경로`);
      render();
    }
    requestAnimationFrame(animateDemo);
  }

  function step(now) {
    if (!running) return;
    const dt = Math.min((now - lastTime) / 1000, .05);
    lastTime = now;
    const fixedStep = 1 / 120;
    physicsAccumulator += dt;
    while (physicsAccumulator >= fixedStep) {
      simulate(fixedStep);
      physicsAccumulator -= fixedStep;
    }
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
    const hitId = a.id === "cue" && b.id !== "cue" ? b.id : (b.id === "cue" && a.id !== "cue" ? a.id : null);
    if (hitId && !shot.touched.has(hitId)) {
      if (shot.touched.size === 0 && hitId !== firstBallId) shot.invalidOrder = true;
      shot.touched.add(hitId);
      const secondId = firstBallId === "yellow" ? "red" : "yellow";
      if (hitId === secondId && shot.touched.has(firstBallId)) {
        shot.secondTouchedAtCushions = shot.cushionCount;
        if (!shot.invalidOrder && shot.secondTouchedAtCushions >= 3) {
          shot.scored = true;
          stats.successes++;
          updateStats();
          setStatus("success", "득점", `${firstBallId === "yellow" ? "노란공" : "빨간공"} 먼저 · 두 번째 적구 전 쿠션 ${shot.secondTouchedAtCushions}회`);
        }
      }
    }
  }

  function finishShot() {
    running = false;
    balls.forEach((b) => { b.vx = 0; b.vy = 0; });
    const success = shot.scored;
    if (success) {
      setStatus("success", "득점", `${firstBallId === "yellow" ? "노란공" : "빨간공"} 먼저 · 두 번째 적구 전 쿠션 ${shot.secondTouchedAtCushions}회`);
    } else {
      const hitText = shot.touched.size === 0 ? "적구 미접촉" : shot.touched.size === 1 ? "적구 1개 접촉" : "두 적구 접촉";
      const reason = shot.invalidOrder ? "선택하지 않은 공을 먼저 맞힘" : `${hitText} · 두 번째 적구 전 쿠션 ${shot.secondTouchedAtCushions ?? shot.cushionCount}회`;
      setStatus("fail", "실패", reason);
    }
    updateStats();
    document.getElementById("retryShotBtn").disabled = !lastShotSetup;
    render();
  }

  function retryLastShot() {
    if (!lastShotSetup || running || solving) return;
    balls = cloneBalls(lastShotSetup.balls);
    successPath = lastShotSetup.successPath;
    angle = lastShotSetup.angle;
    power = lastShotSetup.power;
    sideSpin = lastShotSetup.sideSpin;
    verticalSpin = lastShotSetup.verticalSpin;
    previewCushions = lastShotSetup.previewCushions;
    firstBallId = lastShotSetup.firstBallId;
    demoPosition = null;
    document.getElementById("successShootBtn").disabled = !successPath;
    document.getElementById("retryShotBtn").disabled = true;
    updateFirstHitGuide(successPath);
    document.getElementById("firstYellowBtn").classList.toggle("selected", firstBallId === "yellow");
    document.getElementById("firstRedBtn").classList.toggle("selected", firstBallId === "red");
    document.getElementById("firstYellowBtn").setAttribute("aria-pressed", String(firstBallId === "yellow"));
    document.getElementById("firstRedBtn").setAttribute("aria-pressed", String(firstBallId === "red"));
    updateControls();
    setStatus("", "다시 하기 준비", "샷 실행 전의 공 배치와 경로를 복원했습니다.");
    render();
  }

  function randomize() {
    if (running || solving) return;
    clearSuccessPath();
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
    if (running || solving) return;
    balls = cloneBalls(initial);
    clearSuccessPath();
    angle = 0; power = 55; sideSpin = 0; verticalSpin = 0; previewCushions = 3;
    setStatus("", "설계 중", "공을 드래그하거나 빈 곳을 눌러 조준하세요.");
    updateControls();
    render();
  }

  function selectFirstBall(id) {
    if (running || solving) return;
    firstBallId = id;
    clearSuccessPath();
    const yellowButton = document.getElementById("firstYellowBtn");
    const redButton = document.getElementById("firstRedBtn");
    yellowButton.classList.toggle("selected", id === "yellow");
    redButton.classList.toggle("selected", id === "red");
    yellowButton.setAttribute("aria-pressed", String(id === "yellow"));
    redButton.setAttribute("aria-pressed", String(id === "red"));
    setStatus("", "설계 중", `${id === "yellow" ? "노란공" : "빨간공"}을 먼저 맞히는 경로를 설계합니다.`);
    render();
  }

  function simulateRoute(candidateAngle, candidatePower, candidateSpin, captureFrames = false) {
    const simBalls = cloneBalls(balls);
    const radians = candidateAngle * Math.PI / 180;
    const speed = 290 + candidatePower * 5.1;
    simBalls[0].vx = Math.cos(radians) * speed;
    simBalls[0].vy = Math.sin(radians) * speed;
    const targetSecond = firstBallId === "yellow" ? "red" : "yellow";
    const route = [{ x: simBalls[0].x, y: simBalls[0].y }];
    const events = [];
    let cushions = 0;
    let firstTouched = false;
    let invalid = false;
    let scored = false;
    let scoredCushions = null;
    let lastRecorded = route[0];
    const touching = new Set();
    const dt = 1 / 120;
    const frames = captureFrames ? [simBalls.map(({ x, y, vx, vy }) => ({ x, y, vx, vy }))] : null;

    for (let frame = 0; frame < 1800; frame++) {
      simBalls.forEach((ball) => {
        ball.x += ball.vx * dt;
        ball.y += ball.vy * dt;
        const currentSpeed = Math.hypot(ball.vx, ball.vy);
        const nextSpeed = Math.max(0, currentSpeed - 74 * dt);
        if (currentSpeed) { ball.vx *= nextSpeed / currentSpeed; ball.vy *= nextSpeed / currentSpeed; }
        let bounced = false;
        if (ball.x < bounds.left) { ball.x = bounds.left; ball.vx = Math.abs(ball.vx) * .88; ball.vy += candidateSpin * 22; bounced = true; }
        if (ball.x > bounds.right) { ball.x = bounds.right; ball.vx = -Math.abs(ball.vx) * .88; ball.vy -= candidateSpin * 22; bounced = true; }
        if (ball.y < bounds.top) { ball.y = bounds.top; ball.vy = Math.abs(ball.vy) * .88; ball.vx -= candidateSpin * 22; bounced = true; }
        if (ball.y > bounds.bottom) { ball.y = bounds.bottom; ball.vy = -Math.abs(ball.vy) * .88; ball.vx += candidateSpin * 22; bounced = true; }
        if (bounced && ball.id === "cue") {
          cushions++;
          route.push({ x: ball.x, y: ball.y });
          events.push({ x: ball.x, y: ball.y, kind: "cushion", label: String(cushions) });
        }
      });

      if (!scored && Math.hypot(simBalls[1].x - simBalls[2].x, simBalls[1].y - simBalls[2].y) < ballRadius * 2) {
        invalid = true;
        break;
      }

      for (let i = 0; i < simBalls.length; i++) {
        for (let j = i + 1; j < simBalls.length; j++) {
          const a = simBalls[i], b = simBalls[j];
          const dx = b.x - a.x, dy = b.y - a.y;
          const distance = Math.hypot(dx, dy);
          const pairKey = `${a.id}:${b.id}`;
          if (distance >= ballRadius * 2) { touching.delete(pairKey); continue; }
          if (!distance) continue;
          // 학습용 성공 경로에서는 적구끼리 먼저 충돌해 목표 위치가
          // 바뀌는 우회 해법을 제외하고 수구의 직접 접촉만 인정한다.
          if (!scored && a.id !== "cue" && b.id !== "cue") {
            invalid = true;
            continue;
          }
          const nx = dx / distance, ny = dy / distance;
          const overlap = ballRadius * 2 - distance;
          a.x -= nx * overlap / 2; a.y -= ny * overlap / 2;
          b.x += nx * overlap / 2; b.y += ny * overlap / 2;
          const cueBall = a.id === "cue" ? a : b;
          const targetBall = a.id === "cue" ? b : a;
          const incomingSpeed = Math.hypot(cueBall.vx, cueBall.vy);
          const incomingX = incomingSpeed ? cueBall.vx / incomingSpeed : nx;
          const incomingY = incomingSpeed ? cueBall.vy / incomingSpeed : ny;
          const targetDx = targetBall.x - cueBall.x;
          const targetDy = targetBall.y - cueBall.y;
          const lateralOffset = Math.abs(targetDx * -incomingY + targetDy * incomingX);
          const hitThickness = Math.max(0, Math.min(1, 1 - lateralOffset / (ballRadius * 2)));
          const relative = (a.vx - b.vx) * nx + (a.vy - b.vy) * ny;
          if (relative > 0) {
            const impulse = relative * .96;
            a.vx -= impulse * nx; a.vy -= impulse * ny;
            b.vx += impulse * nx; b.vy += impulse * ny;
          }
          if (!scored && !touching.has(pairKey) && (a.id === "cue" || b.id === "cue")) {
            const hitId = a.id === "cue" ? b.id : a.id;
            const cue = simBalls[0];
            route.push({ x: cue.x, y: cue.y });
            const contactDistance = Math.hypot(cue.x - targetBall.x, cue.y - targetBall.y) || 1;
            events.push({
              x: cue.x,
              y: cue.y,
              kind: "ball",
              label: hitId === "yellow" ? "Y" : "R",
              thickness: hitThickness,
              contactX: (cue.x - targetBall.x) / contactDistance,
              contactY: (cue.y - targetBall.y) / contactDistance,
            });
            if (!firstTouched) {
              if (hitId !== firstBallId) invalid = true;
              else firstTouched = true;
            } else if (hitId === targetSecond && cushions >= 3 && !invalid) {
              route.push({ x: cue.x, y: cue.y });
              if (frames) frames.push(simBalls.map(({ x, y, vx, vy }) => ({ x, y, vx, vy })));
              if (!captureFrames) return { angle: candidateAngle, power: candidatePower, spin: candidateSpin, cushions, points: route, events, frames };
              scored = true;
              scoredCushions = cushions;
            }
          }
          touching.add(pairKey);
        }
      }

      const cue = simBalls[0];
      if (Math.hypot(cue.x - lastRecorded.x, cue.y - lastRecorded.y) > 24) {
        lastRecorded = { x: cue.x, y: cue.y };
        route.push(lastRecorded);
      }
      if (frames && frame % 2 === 0) frames.push(simBalls.map(({ x, y, vx, vy }) => ({ x, y, vx, vy })));
      if (invalid || simBalls.every((ball) => Math.hypot(ball.vx, ball.vy) < 5)) break;
    }
    return scored
      ? { angle: candidateAngle, power: candidatePower, spin: candidateSpin, cushions: scoredCushions, points: route, events, frames }
      : null;
  }

  function angleDistance(a, b) {
    return Math.abs(((a - b + 540) % 360) - 180);
  }

  function solveSuccessRoute() {
    if (running) return;
    const solveButton = document.getElementById("solveBtn");
    if (solving) {
      solving = false;
      solveRunId++;
      solveButton.textContent = "3쿠션 성공 경로 찾기";
      setStatus("", "탐색 취소", "경로 계산을 중단했습니다. 공 위치와 조건을 다시 조정할 수 있습니다.");
      return;
    }
    solving = true;
    const currentRunId = ++solveRunId;
    clearSuccessPath();
    solveButton.textContent = "탐색 취소";
    setStatus("", "계산 중", `${firstBallId === "yellow" ? "노란공" : "빨간공"}을 먼저 맞히는 3쿠션 경로를 탐색합니다.`);
    render();

    const cue = balls[0];
    const first = balls.find((ball) => ball.id === firstBallId);
    const directAngle = Math.atan2(first.y - cue.y, first.x - cue.x) * 180 / Math.PI;
    const angles = Array.from({ length: 360 }, (_, index) => index - 180).sort((a, b) => angleDistance(a, directAngle) - angleDistance(b, directAngle));
    const powers = [72, 88, 58, 100, 44];
    const spins = [0, -.5, .5, -1, 1];
    let angleIndex = 0;
    let powerIndex = 0;
    let spinIndex = 0;
    let testedCount = 0;
    const totalCandidates = angles.length * powers.length * spins.length;
    let found = null;

    function searchBatch() {
      if (!solving || currentRunId !== solveRunId) return;
      const batchStartedAt = performance.now();
      while (!found && angleIndex < angles.length && performance.now() - batchStartedAt < 10) {
        found = simulateRoute(angles[angleIndex], powers[powerIndex], spins[spinIndex]);
        testedCount++;
        spinIndex++;
        if (spinIndex >= spins.length) {
          spinIndex = 0;
          powerIndex++;
        }
        if (powerIndex >= powers.length) {
          powerIndex = 0;
          angleIndex++;
        }
      }
      if (!found && angleIndex < angles.length) {
        document.getElementById("statusText").textContent = `성공 경로 탐색 ${Math.round(testedCount / totalCandidates * 100)}% · 필요하면 탐색을 취소할 수 있습니다.`;
        setTimeout(searchBatch, 0);
        return;
      }
      solving = false;
      solveButton.textContent = "3쿠션 성공 경로 찾기";
      if (found) {
        found = simulateRoute(found.angle, found.power, found.spin, true) || found;
        successPath = found;
        updateFirstHitGuide(found);
        angle = found.angle;
        power = found.power;
        sideSpin = found.spin;
        verticalSpin = 0;
        previewCushions = Math.min(5, Math.max(3, found.cushions));
        updateControls();
        document.getElementById("successShootBtn").disabled = false;
        setStatus("success", "경로 발견", `${firstBallId === "yellow" ? "노란공" : "빨간공"} 먼저 · ${found.cushions}쿠션 · 각도 ${Math.round(found.angle)}° · 세기 ${found.power}%`);
      } else {
        document.getElementById("successShootBtn").disabled = true;
        setStatus("fail", "경로 없음", "현재 배치에서는 계산 범위 안의 성공 경로를 찾지 못했습니다. 공 위치를 조금 바꾸거나 다시 시도하세요.");
      }
      render();
    }
    setTimeout(searchBatch, 30);
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
    const layout = { id: String(Date.now()), name, balls: balls.map(({ x, y }) => ({ x, y })), angle, power, sideSpin, verticalSpin, previewCushions, firstBallId };
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
    selectFirstBall(layout.firstBallId || "yellow");
    clearSuccessPath();
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
  document.getElementById("successShootBtn").addEventListener("click", playSuccessRoute);
  document.getElementById("retryShotBtn").addEventListener("click", retryLastShot);
  document.getElementById("solveBtn").addEventListener("click", solveSuccessRoute);
  document.getElementById("firstYellowBtn").addEventListener("click", () => selectFirstBall("yellow"));
  document.getElementById("firstRedBtn").addEventListener("click", () => selectFirstBall("red"));
  document.getElementById("randomBtn").addEventListener("click", randomize);
  document.getElementById("resetBtn").addEventListener("click", reset);
  document.getElementById("clearStatsBtn").addEventListener("click", () => { stats.attempts = 0; stats.successes = 0; updateStats(); });
  window.addEventListener("resize", render);

  updateControls();
  updateStats();
  refreshLayouts();
  render();
})();
