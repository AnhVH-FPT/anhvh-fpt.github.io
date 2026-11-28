/* Contact section (index.html and research.html): small copy-to-clipboard icon buttons */
document.querySelectorAll('.ccopy').forEach(b => b.addEventListener('click', async () => {
  const text = b.dataset.copy;
  let ok = false;
  try { await navigator.clipboard.writeText(text); ok = true; }
  catch {
    const ta = Object.assign(document.createElement('textarea'), { value: text });
    document.body.appendChild(ta); ta.select();
    ok = !!(document.execCommand && document.execCommand('copy'));
    document.body.removeChild(ta);
  }
  if (!ok) {                                   // clipboard blocked: select the value so Ctrl+C works
    const el = b.parentElement.querySelector('a') || b.parentElement;
    const r = document.createRange(); r.selectNodeContents(el);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    b.title = 'Press Ctrl+C to copy'; return;
  }
  b.classList.add('done'); b.setAttribute('aria-label', 'Copied');
  setTimeout(() => { b.classList.remove('done'); b.setAttribute('aria-label', b.title); }, 1600);
}));

/* Footer visit counter: total page views from GoatCounter's public counter
   (needs "Allow adding visitor counts on your website" in the GoatCounter site settings).
   If the number cannot be read (setting off, blocked, offline), the box simply stays hidden. */
(() => {
  const box = document.getElementById('visits'), out = document.getElementById('visits-n');
  if (!box || !out || !window.fetch) return;
  const ctl = 'AbortController' in window ? new AbortController() : null;
  if (ctl) setTimeout(() => ctl.abort(), 6000);
  fetch('https://haanhvu010799.goatcounter.com/counter/TOTAL.json', ctl ? { signal: ctl.signal } : {})
    .then(r => (r.ok ? r.json() : Promise.reject(r.status)))
    .then(d => {
      const n = Number(String(d && d.count != null ? d.count : '').replace(/[^0-9]/g, ''));
      if (!n) return;
      out.textContent = n.toLocaleString('en-US');
      box.hidden = false;
    })
    .catch(() => { /* keep the box hidden */ });
})();
