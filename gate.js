/* 비밀번호 확인 — 맞는 비밀번호로만 기록 파일(parkville-data.enc)의 암호가 풀린다.
   비밀번호는 어디에도 저장·전송하지 않고, 암호 풀기는 이 브라우저 안에서만 한다. */
(async function () {
  const load = (src) => new Promise((ok, no) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = no; document.body.appendChild(s); });
  const boot = async () => { document.body.classList.remove('gated'); await load('excel.js?v=9'); await load('script.js?v=9'); };
  const b64 = (t) => Uint8Array.from(atob(t), (c) => c.charCodeAt(0));
  let enc;
  try { const r = await fetch('parkville-data.enc', { cache: 'no-store' }); if (!r.ok) throw 0; enc = await r.json(); } catch (e) { return boot(); }
  const gate = document.getElementById('gate'), pw = document.getElementById('gate-pw'), msg = document.getElementById('gate-msg'), btn = document.getElementById('gate-btn');
  gate.hidden = false; pw.focus();
  gate.addEventListener('submit', async (e) => {
    e.preventDefault(); if (!pw.value) return;
    btn.disabled = true; msg.textContent = '확인 중…'; msg.className = '';
    try {
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw.value), 'PBKDF2', false, ['deriveKey']);
      const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: b64(enc.salt), iterations: enc.iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
      const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64(enc.iv) }, key, b64(enc.ct));
      window.PARKVILLE_DATA = JSON.parse(new TextDecoder().decode(plain));
      pw.value = ''; gate.hidden = true; boot();
    } catch (err) { msg.textContent = '비밀번호가 맞지 않습니다.'; msg.className = 'bad'; btn.disabled = false; pw.select(); }
  });
})();
