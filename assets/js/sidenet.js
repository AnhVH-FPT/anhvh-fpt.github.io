/* Neural-network animation in the empty left and right margins (index.html and research.html).
   A fixed canvas over the page that only draws outside the content column, so text is never covered. */
(() => {
  const root = document.documentElement;
  const still = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const cv = Object.assign(document.createElement('canvas'), { className: 'sidenet' });
  cv.setAttribute('aria-hidden', 'true');
  document.body.appendChild(cv);
  const g = cv.getContext('2d');
  const CONTENT = 1040, GAP = 12, MIN = 90;          // .wrap max-width; keep 12px off the content; hide if a margin is narrower
  let W = 0, H = 0, side = 0, nodes = [], mx = -999, my = -999, dpr = 1;
  const fit = () => {
    dpr = Math.min(devicePixelRatio || 1, 2);
    W = innerWidth; H = innerHeight;
    cv.width = W * dpr; cv.height = H * dpr; g.setTransform(dpr, 0, 0, dpr, 0, 0);
    side = (W - CONTENT) / 2 - GAP;
    nodes = [];
    if (side < MIN) { g.clearRect(0, 0, W, H); return; }
    const n = Math.max(6, Math.round(side * H / 14000));
    for (const left of [true, false])
      for (let i = 0; i < n; i++) {
        const x0 = left ? 0 : W - side;
        nodes.push({ x0, x: x0 + Math.random() * side, y: Math.random() * H, vx: (Math.random() - .5) * .3, vy: (Math.random() - .5) * .3 });
      }
    if (still) tick();
  };
  const line = (x1, y1, x2, y2, a) => { g.globalAlpha = a; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke(); };
  function tick() {
    g.clearRect(0, 0, W, H);
    if (nodes.length) {
      g.strokeStyle = g.fillStyle = getComputedStyle(root).getPropertyValue('--net').trim();
      nodes.forEach((a, i) => {
        if (!still) {
          a.x += a.vx; a.y += a.vy;
          if (a.x < a.x0 || a.x > a.x0 + side) a.vx *= -1;
          if (a.y < 0 || a.y > H) a.vy *= -1;
        }
        g.globalAlpha = .55; g.beginPath(); g.arc(a.x, a.y, 2.2, 0, 6.3); g.fill();
        for (let j = i + 1; j < nodes.length; j++) {
          const b = nodes[j]; if (b.x0 !== a.x0) continue;
          const d = Math.hypot(a.x - b.x, a.y - b.y);
          if (d < 130) line(a.x, a.y, b.x, b.y, (1 - d / 130) * .35);
        }
        const dm = Math.hypot(a.x - mx, a.y - my);
        if (dm < 160) line(a.x, a.y, mx, my, (1 - dm / 160) * .7);
      });
      g.globalAlpha = 1;
      const hero = document.querySelector('.hero');   // the home hero has its own full-width network
      if (hero) { const r = hero.getBoundingClientRect(); if (r.bottom > 0) g.clearRect(0, r.top, W, r.height); }
    }
    if (!still) requestAnimationFrame(tick);
  }
  addEventListener('pointermove', e => { mx = e.clientX; my = e.clientY; }, { passive: true });
  document.addEventListener('pointerleave', () => { mx = my = -999; });
  addEventListener('resize', fit);
  fit();
  if (!still) tick();
})();
