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
