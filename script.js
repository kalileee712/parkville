/* 파크빌 관리 — 전기검침 · 임대료 · 관리비 */
(function () {
  const { workbook, ST, colName } = window.XLSXLite;
  const LS_KEY = 'parkville-v3';
  const FILE_NAME = '파크빌_관리대장.xlsx';
  const LATE_RATE = 0.05; // 미납 시 다음 달 가산료 5%
  const TEST = location.hash.startsWith('#test');

  /* ---------- 기본 자료 (이전 관리 파일에서 정리) ---------- */
  // 2~5층 고정비 내역. 1층(101/102, 103, 104)은 일반관리비만 고정.
  const FIXED_ITEMS = [['일반관리비', 60000], ['전기 기본요금', 22000], ['TV 수신료', 2500], ['인터넷·유선요금', 11500]];
  // 호실 구성만 여기 두고, 이름·임대료·지난 기록은 자료 파일(parkville-data.js)에서 가져온다
  const BASE_ROOMS = [['101', 1, '101/102호'], ['103', 1], ['104', 1], ...[2, 3, 4, 5].flatMap((f) => [1, 2, 3, 4, 5, 6, 7].map((i) => [String(f * 100 + i), f]))]
    .map(([id, floor, label]) => ({ id, floor, rent: 0, tenant: '', label: label || id + '호', fixed: 0 }));
  const METERED = BASE_ROOMS.filter((r) => r.floor > 1).map((r) => r.id);
  const FIRST_PAY_MONTH = '2026-11';
  const newMonth = (prev) => ({ kepco: null, fixedItems: JSON.parse(JSON.stringify(prev ? prev.fixedItems : FIXED_ITEMS)), r: {} });

  // 자료 파일에 나중에 추가된 수정분(patches)을 이미 쓰고 있는 기록에도 한 번씩만 반영한다
  function applyPatches(s, data) {
    if (!data || !data.patches) return;
    if (!s.applied) s.applied = [];
    data.patches.forEach((p) => {
      if (s.applied.includes(p.id)) return;
      const m = s.months[p.ym] || (s.months[p.ym] = newMonth());
      for (const id in p.usage || {}) (m.r[id] || (m.r[id] = {})).usage = p.usage[id];
      s.applied.push(p.id);
    });
  }
  function seed(data = window.PARKVILLE_DATA || null) {
    const now = new Date(), thisYm = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
    const s = { rooms: JSON.parse(JSON.stringify(data ? data.rooms : BASE_ROOMS)), months: {}, cur: data ? FIRST_PAY_MONTH : thisYm, tab: 'meter', savedAt: null, fileName: null,
      hasData: !!data, cfg: { ...CFG, ...(data && data.cfg) }, sel: '201' };
    if (!data) return s;
    for (const ym in data.reads || {}) {
      const m = s.months[ym] || (s.months[ym] = newMonth());
      for (const id in data.reads[ym]) m.r[id] = { read: data.reads[ym][id] };
    }
    // 이전 관리 파일 기록: 이름·임대료·관리비·입금 여부·한전요금. 지난 달은 저장된 것으로 보고 잠근다.
    const hist = data.months || {};
    for (const ym in hist) {
      const m = s.months[ym] || (s.months[ym] = newMonth());
      if (hist[ym].kepco != null) m.kepco = hist[ym].kepco;
      for (const id in hist[ym].r) m.r[id] = { ...m.r[id], ...hist[ym].r[id] };
    }
    for (const ym in s.months) if (ym < FIRST_PAY_MONTH) s.months[ym].lock = { meter: true, pay: true };
    applyPatches(s, data);
    return s;
  }
  const CFG = { bank: '', holder: '', dueDay: 25, notice: '월세는 매월입주일 기준 선불로 입금 요망함' };
  let S;
  try { S = (!TEST && JSON.parse(localStorage.getItem(LS_KEY))) || seed(); } catch (e) { S = seed(); }
  if (!S.cfg) S.cfg = { ...CFG };
  if (S.sel === undefined) S.sel = '201';
  if (S.hasData === undefined) S.hasData = true;
  applyPatches(S, window.PARKVILLE_DATA);
  /* ---------- 되돌리기 ---------- */
  // 내용이 바뀔 때마다 바뀌기 전 상태를 쌓아 두었다가 「되돌리기」를 누르면 하나씩 되돌린다 (달·탭 이동은 제외)
  const dataJson = () => JSON.stringify({ rooms: S.rooms, cfg: S.cfg, months: Object.fromEntries(Object.entries(S.months).filter(([, m]) => Object.keys(m.r).length || m.kepco != null || m.kepcoF || m.common != null || m.lock)) });
  const undoStack = [];
  let lastJson = dataJson(), lastView = { cur: S.cur, tab: S.tab };
  function track() {
    const now = dataJson();
    if (now !== lastJson) { undoStack.push({ json: lastJson, ...lastView }); if (undoStack.length > 50) undoStack.shift(); lastJson = now; }
    lastView = { cur: S.cur, tab: S.tab };
    const b = document.getElementById('undobtn'); if (b) b.disabled = !undoStack.length;
  }
  function undo() {
    const u = undoStack.pop(); if (!u) return;
    const d = JSON.parse(u.json);
    S.rooms = d.rooms; S.cfg = d.cfg; S.months = d.months; S.cur = u.cur; S.tab = u.tab; lastJson = u.json;
    render(); toast('바로 전 상태로 되돌렸습니다.');
  }
  const persist = () => { track(); if (TEST) return; try { localStorage.setItem(LS_KEY, JSON.stringify(S)); } catch (e) { } };

  /* ---------- 계산 ---------- */
  const shift = (ym, d) => { let [y, m] = ym.split('-').map(Number); m += d; while (m < 1) { m += 12; y--; } while (m > 12) { m -= 12; y++; } return y + '-' + String(m).padStart(2, '0'); };
  const ymLong = (ym) => { const [y, m] = ym.split('-'); return `${y}년 ${+m}월`; };
  const ymShort = (ym) => { const [y, m] = ym.split('-'); return `${y.slice(2)}.${+m}월`; };
  const won = (n) => (n == null || n === '' || isNaN(n) ? '' : Math.round(n).toLocaleString('ko-KR'));
  const num = (v) => { const t = String(v).replace(/[^\d]/g, ''); return t === '' ? null : Number(t); };
  const h = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

  const month = (ym) => S.months[ym] || (S.months[ym] = newMonth(S.months[shift(ym, -1)] || S.months[Object.keys(S.months).filter((k) => k < ym).sort().pop()]));
  const M = (ym) => S.months[ym] || { kepco: null, fixedItems: FIXED_ITEMS, r: {} };
  const rec = (ym, id) => { const m = month(ym); return m.r[id] || (m.r[id] = {}); };
  const peek = (ym, id) => M(ym).r[id] || {};
  // 이름·임대료처럼 한 번 넣으면 다음 달에도 이어지는 값: 그 달에 없으면 가장 가까운 이전 달 값을 쓴다
  function inherit(room, ym, f) {
    const keys = Object.keys(S.months).filter((k) => k <= ym).sort().reverse();
    for (const k of keys) { const v = peek(k, room.id)[f]; if (v != null) return v; }
    return room[f];
  }
  // 사용량: 직접 넣은 값이 있으면 그 값, 없으면 당월 검침 − 전월 검침
  const usageOf = (room, ym) => { const c = peek(ym, room.id), a = c.read, b = peek(shift(ym, -1), room.id).read; return room.floor <= 1 ? null : c.usage != null ? c.usage : a != null && b != null ? a - b : null; };
  const floorUsage = (ym, f) => byFloor(f).reduce((s, r) => { const u = usageOf(r, ym); return s + (u > 0 ? u : 0); }, 0);
  const floorKepco = (ym, f) => { const k = M(ym).kepcoF; return k && k[f] != null ? k[f] : null; };
  // 그 달 한전 요금 합계: 층별 요금 + 공용 (이전 기록은 총액만 있음)
  const kepcoSum = (ym) => { const m = M(ym), fs = [2, 3, 4, 5].map((f) => floorKepco(ym, f)).filter((v) => v != null); return fs.length ? fs.reduce((a, b) => a + b, 0) + (m.common || 0) : m.kepco; };
  const totalUsage = (ym) => S.rooms.reduce((s, r) => { const u = usageOf(r, ym); return s + (u > 0 ? u : 0); }, 0);
  const fixedSum = (ym) => M(ym).fixedItems.reduce((s, x) => s + (x[1] || 0), 0);

  function calc(room, ym = S.cur) {
    const c = peek(ym, room.id), metered = room.floor > 1;
    const usage = usageOf(room, ym), tot = metered ? floorUsage(ym, room.floor) : 0, kepco = floorKepco(ym, room.floor);
    // 전기 사용료 = 그 층 한전 요금 × (이 호실 사용량 ÷ 그 층 전체 사용량) + 공동전기료
    // 공용 전기요금은 건물 전체 1건 — 2~5층 호실 수로 똑같이 나눈다(공동전기료)
    const common = M(ym).common, commonShare = metered && common != null ? Math.round(common / METERED.length) : null;
    const autoFee = metered && usage != null && usage >= 0 && kepco != null && tot > 0 ? Math.round((kepco * usage) / tot) + (commonShare || 0) : null;
    let fee = c.fee != null ? c.fee : autoFee;
    const rent = inherit(room, ym, 'rent');
    let fixed = metered ? fixedSum(ym) : inherit(room, ym, 'fixed');
    if (c.mgmtAmt != null) { fixed = Math.min(fixed, c.mgmtAmt); fee = c.mgmtAmt - fixed; } // 이전 파일에서 옮긴 관리비 총액
    return { commonShare, prev: peek(shift(ym, -1), room.id).read, read: c.read, usage, autoFee, fee, feeManual: c.fee != null, rent, fixed, mgmt: fixed + (fee || 0), tenant: inherit(room, ym, 'tenant') || '', rentPaid: c.rentPaid || '', mgmtPaid: c.mgmtPaid || '', note: c.note || '', metered, bad: usage != null && usage < 0 };
  }
  // 이전 달들에 미납으로 남아 있는 금액과 가산료(5%)
  function arrears(room, ym = S.cur) {
    const items = [];
    Object.keys(S.months).filter((k) => k < ym && k >= FIRST_PAY_MONTH).sort().forEach((k) => {
      const c = calc(room, k);
      if (c.rentPaid === 'N') items.push({ ym: k, f: 'rentPaid', what: '임대료', amt: c.rent });
      if (c.mgmtPaid === 'N') items.push({ ym: k, f: 'mgmtPaid', what: '관리비', amt: c.mgmt });
    });
    const principal = items.reduce((s, x) => s + x.amt, 0);
    return { items, principal, late: Math.round(principal * LATE_RATE) };
  }
  // 납부영수증 한 장에 들어갈 숫자
  function receipt(room, ym = S.cur) {
    const c = calc(room, ym), m = M(ym), it = m.fixedItems, [y, mo] = ym.split('-').map(Number), p = shift(ym, -1);
    const split = c.metered && c.fixed === fixedSum(ym);
    const general = split ? it[0][1] : c.fixed, base = split ? it[1][1] : null, tv = split ? it[2][1] : null, net = split ? it[3][1] : null;
    const fee = c.fee || 0, common = c.metered && c.commonShare && fee ? Math.min(c.commonShare, fee) : 0, use = fee - common;
    const un = arrears(room, ym).items.filter((x) => x.f === 'mgmtPaid'), prin = un.reduce((s, x) => s + x.amt, 0), unpaid = prin + Math.round(prin * LATE_RATE);
    const total = c.mgmt, late = Math.round(total * LATE_RATE);
    const nm = (c.tenant || '').split(/[(.]/).filter(Boolean).pop() || '', name = nm.length === 3 ? nm[0] + '*' + nm[2] : nm.length === 2 ? nm[0] + '*' : nm;
    return { room, c, name, general, base, tv, net, use, common, usage: c.usage, elecSub: (base || 0) + use + common + (tv || 0), total, unpaid,
      unpaidLabel: un.length ? [...new Set(un.map((x) => +x.ym.slice(5) + '월'))].join('·') : `${+p.slice(5)}월`, within: total + unpaid, late, after: total + unpaid + late,
      period: `${p.slice(2, 4)}.${p.slice(5)}/10~${String(y).slice(2)}.${String(mo).padStart(2, '0')}/09`, y, mo, due: `${String(y).slice(2)}년 ${mo}월${S.cfg.dueDay}일` };
  }
  function summary(ym = S.cur) {
    const t = { readDone: 0, readAll: METERED.length, rentY: 0, rentN: 0, mgmtY: 0, mgmtN: 0, rentSum: 0, mgmtSum: 0, fixedSum: 0, feeSum: 0, unpaid: 0, unpaidAmt: 0, arrP: 0, arrL: 0, arrN: 0 };
    S.rooms.forEach((room) => {
      const c = calc(room, ym), a = arrears(room, ym);
      if (c.metered && c.usage != null) t.readDone++;
      t.rentSum += c.rent; t.mgmtSum += c.mgmt; t.fixedSum += c.fixed; t.feeSum += c.fee || 0;
      if (c.rentPaid === 'Y') t.rentY++; else if (c.rentPaid === 'N') { t.rentN++; t.unpaidAmt += c.rent; }
      if (c.mgmtPaid === 'Y') t.mgmtY++; else if (c.mgmtPaid === 'N') { t.mgmtN++; t.unpaidAmt += c.mgmt; }
      if (c.rentPaid === 'N' || c.mgmtPaid === 'N') t.unpaid++;
      if (a.principal) { t.arrN++; t.arrP += a.principal; t.arrL += a.late; }
    });
    return t;
  }
  const FLOORS = [1, 2, 3, 4, 5];
  const byFloor = (f) => S.rooms.filter((r) => r.floor === f);

  /* ---------- 화면 부품 ---------- */
  const paidText = { Y: '완납', N: '미납', '': '' };
  const seg = (id, f, v) => `<span class="seg" role="group">
    <button type="button" class="ok ${v === 'Y' ? 'on' : ''}" data-act="pay" data-room="${id}" data-f="${f}" data-v="Y" aria-pressed="${v === 'Y'}">완납</button>
    <button type="button" class="no ${v === 'N' ? 'on' : ''}" data-act="pay" data-room="${id}" data-f="${f}" data-v="N" aria-pressed="${v === 'N'}">미납</button></span>`;
  const inp = (id, f, val, o = {}) => `<input class="in ${o.cls || ''}" type="text" ${o.text ? '' : 'inputmode="numeric"'} data-room="${id}" data-f="${f}" data-k="${f}|${id}" value="${h(o.text ? val : won(val))}" placeholder="${h(o.ph || '')}" aria-label="${h(o.label || f)}" autocomplete="off" ${o.ro ? 'readonly tabindex="-1"' : ''}>`;
  const usageCell = (c) => c.usage == null ? '<span class="muted">–</span>' : c.bad ? `<span class="warn">전월보다 작음</span>` : `${won(c.usage)}`;

  function kpis() {
    const t = summary(), pct = (a, b) => (b ? Math.round((a / b) * 100) : 0), n = S.rooms.length;
    return `<section class="kpis" aria-label="이번 달 요약">
      <div class="kpi"><span>전기 검침</span><b>${t.readDone}<i>/${t.readAll}호</i></b><div class="bar"><u style="width:${pct(t.readDone, t.readAll)}%"></u></div></div>
      <div class="kpi"><span>임대료 완납</span><b>${t.rentY}<i>/${n}호</i></b><div class="bar"><u style="width:${pct(t.rentY, n)}%"></u></div></div>
      <div class="kpi"><span>관리비 완납</span><b>${t.mgmtY}<i>/${n}호</i></b><div class="bar"><u style="width:${pct(t.mgmtY, n)}%"></u></div></div>
      <div class="kpi ${t.unpaid ? 'alert' : ''}"><span>이번 달 미납</span><b>${t.unpaid}<i>호 · ${won(t.unpaidAmt) || 0}원</i></b></div>
      <div class="kpi ${t.arrN ? 'alert' : ''}"><span>전월 미납 + 가산료 5%</span><b>${t.arrN}<i>호 · ${won(t.arrP + t.arrL) || 0}원</i></b></div>
    </section>`;
  }

  function viewMeter() {
    const m = month(S.cur), tot = totalUsage(S.cur), ks = kepcoSum(S.cur);
    const unitOf = (f) => { const k = floorKepco(S.cur, f), u = floorUsage(S.cur, f); return k != null && u > 0 ? (k / u).toLocaleString('ko-KR', { maximumFractionDigits: 1 }) : null; };
    let rows = '', alloc = 0;
    FLOORS.forEach((f) => {
      const k = floorKepco(S.cur, f), info = f === 1 ? '' : `<span class="fl-info">한전 요금 <b>${k != null ? won(k) + '원' : '미입력'}</b> · 사용량 ${won(floorUsage(S.cur, f)) || 0} kWh${unitOf(f) ? ` · 1kWh당 ${unitOf(f)}원` : ''}</span>`;
      rows += `<tr class="floor"><td colspan="6">${f}층${info}</td></tr>`;
      byFloor(f).forEach((room) => {
        const c = calc(room); alloc += c.autoFee || 0;
        const name = `<th scope="row">${room.label}</th><td>${inp(room.id, 'tenant', c.tenant, { text: 1, cls: 'note', ph: '이름(공실이면 비움)', label: room.label + ' 입주자 이름' })}</td>`;
        if (!c.metered) { rows += `<tr data-row="${room.id}">${name}<td colspan="4" class="muted c">상가 — 전기요금은 입주자가 직접 납부 (입력할 것 없음)</td></tr>`; return; }
        rows += `<tr data-row="${room.id}">${name}<td class="n">${won(c.prev)}</td>
          <td>${inp(room.id, 'read', c.read, { label: room.label + ' 당월 검침' })}</td><td>${inp(room.id, 'usage', c.usage, { cls: 'big' + (c.bad ? ' bad' : ''), ph: '사용량', label: room.label + ' 사용량 kWh' })}</td><td class="n strong">${c.autoFee == null ? '<span class="muted">–</span>' : won(c.autoFee)}</td></tr>`;
      });
    });
    return `<div class="sheet"><div class="kepco">
        <div class="ktitle"><b>한전 전기요금 (층별)</b><small>한전에 전화해서 확인한 이번 달 요금을 층마다 넣습니다</small></div>
        ${[2, 3, 4, 5].map((f) => `<label class="kf">${f}층 <input class="in big" type="text" inputmode="numeric" data-f="kepcoF" data-floor="${f}" data-k="kepcoF|${f}" value="${won(floorKepco(S.cur, f))}" placeholder="${f}층 요금" autocomplete="off" aria-label="${f}층 한전 전기요금"> 원</label>`).join('')}
        <label class="kf common">공용 <input id="common" class="in big" type="text" inputmode="numeric" data-f="common" data-k="common" value="${won(m.common)}" placeholder="공용 요금" autocomplete="off" aria-label="공용 전기요금"> 원<small>건물 전체 1건 — ${METERED.length}개 호실에 똑같이 나눠 「공동전기료」로</small></label>
        <dl><div><dt>전체 사용량</dt><dd>${won(tot) || 0} kWh</dd></div><div><dt>한전 요금 합계</dt><dd>${ks != null ? won(ks) + ' 원' : '–'}</dd></div><div><dt>호실 배분 합계</dt><dd>${alloc ? won(alloc) + ' 원' : '–'}</dd></div></dl>
      </div>
      <table class="grid"><caption>${ymShort(S.cur)} 전기검침 · 검침일 ${+S.cur.slice(5)}월 10일</caption>
      <thead><tr><th>호수</th><th>이름</th><th>전월 검침</th><th>당월 검침</th><th>사용량(kWh)</th><th>전기 사용료(원)</th></tr></thead><tbody>${rows}</tbody></table></div>
      <p class="hint">① 호실별 <b>사용량(kWh)</b> 을 넣고 Enter 를 누르면 다음 호실로 넘어갑니다. 당월 검침 숫자를 넣으면 사용량이 자동으로 계산되고, 사용량을 넣으면 당월 검침이 맞춰집니다. 이름은 이 탭에서만 고칠 수 있습니다. ② 층별 한전 요금을 넣으면 <b>그 층 요금 × (호실 사용량 ÷ 그 층 사용량)</b> 에 공동전기료를 더해 호실별 전기 사용료가 계산되고, 납부현황의 사용료 칸에 들어갑니다.</p>`;
  }

  function viewPay() {
    const m = month(S.cur), t = summary();
    let rows = '';
    FLOORS.forEach((f) => {
      rows += `<tr class="floor"><td colspan="9">${f}층</td></tr>`;
      byFloor(f).forEach((room) => {
        const c = calc(room), a = arrears(room);
        const arr = a.principal ? `<div class="arr"><b>${won(a.principal + a.late)}</b><small>${a.items.map((x) => `${ymShort(x.ym)} ${x.what} ${won(x.amt)}`).join(', ')} + 가산료 ${won(a.late)}</small>
          <span class="arr-btns"><button type="button" class="mini-btn ok" data-act="settle" data-room="${room.id}" data-v="in" title="기간 안에 받았으면 가산료 없이 완납 처리">기간 내 받음</button><button type="button" class="mini-btn" data-act="settle" data-room="${room.id}" data-v="late" title="가산료까지 받았으면 완납 처리하고 비고에 기록">가산료 포함 받음</button></span></div>` : '';
        rows += `<tr data-row="${room.id}" class="${c.rentPaid === 'N' || c.mgmtPaid === 'N' || a.principal ? 'row-bad' : ''}"><th scope="row">${room.label}<small>${h(c.tenant) || '공실'}</small></th>
          <td>${inp(room.id, 'rent', c.rent, { label: room.label + ' 임대료' })}</td><td class="c">${seg(room.id, 'rentPaid', c.rentPaid)}</td>
          <td>${c.metered ? `<span class="fixed" title="${h(m.fixedItems.map((x) => x[0] + ' ' + won(x[1])).join(' + '))}">${won(c.fixed)}</span>` : inp(room.id, 'fixed', c.fixed, { label: room.label + ' 고정 관리비' })}</td>
          <td>${c.metered ? inp(room.id, 'fee', c.fee, { ph: '검침·한전요금 입력 후 자동', cls: c.feeManual ? 'manual' : '', label: room.label + ' 사용료' }) : '<span class="fixed muted" title="1층 상가는 전기요금을 직접 납부합니다">–</span>'}</td>
          <td class="n strong">${won(c.mgmt)}</td><td class="c">${seg(room.id, 'mgmtPaid', c.mgmtPaid)}</td>
          <td class="arr-cell">${arr}</td>
          <td>${inp(room.id, 'note', c.note, { text: 1, cls: 'note', label: room.label + ' 비고' })}</td></tr>`;
      });
    });
    const fixedPanel = `<details class="fixedbox"><summary>2~5층 고정비 내역 <b>${won(fixedSum(S.cur))}원</b> <span class="muted">(${m.fixedItems.map((x) => x[0] + ' ' + won(x[1])).join(' + ')})</span> — 눌러서 금액 수정</summary>
      <div class="fixeditems">${m.fixedItems.map((x, i) => `<label>${x[0]} <input class="in" type="text" inputmode="numeric" data-f="fixedItem" data-i="${i}" data-k="fixedItem|${i}" value="${won(x[1])}"> 원</label>`).join('')}
      <p class="muted">1층은 일반관리비만 고정입니다 (101/102호 100,000원 · 103호 60,000원 · 104호 60,000원). 여기서 바꾼 금액은 이번 달부터 적용됩니다.</p></div></details>`;
    return fixedPanel + `<div class="sheet"><table class="grid pay"><caption>${ymShort(S.cur)} 월별 임대료·관리비 납부현황</caption>
      <thead><tr><th rowspan="2">호수 · 이름</th><th rowspan="2">임대료</th><th rowspan="2">완납/미납</th><th colspan="3">관리비</th><th rowspan="2">완납/미납</th><th rowspan="2">전월 미납<br>+ 가산료 5%</th><th rowspan="2">비고</th></tr>
      <tr><th>고정</th><th>사용료</th><th>계</th></tr></thead><tbody>${rows}</tbody>
      <tfoot><tr><th>합계</th><td class="n">${won(t.rentSum)}</td><td class="c">완납 ${t.rentY} · 미납 ${t.rentN}</td><td class="n">${won(t.fixedSum)}</td><td class="n">${won(t.feeSum)}</td><td class="n">${won(t.mgmtSum)}</td><td class="c">완납 ${t.mgmtY} · 미납 ${t.mgmtN}</td><td class="n">${won(t.arrP + t.arrL)}</td><td></td></tr></tfoot></table></div>
      <p class="hint">미납으로 표시한 금액은 다음 달 표의 「전월 미납 + 가산료 5%」 칸에 자동으로 나타납니다. 돈을 받으면 그 칸의 버튼을 누르거나, 해당 달로 돌아가 완납으로 바꾸면 사라집니다. 이름은 「① 전기 검침」 탭에서 고칩니다.</p>`;
  }

  /* ---------- 저장 후 잠금 ---------- */
  const TAB_NAME = { meter: '전기 검침', pay: '임대료·관리비 납부현황' };
  const locked = (ym = S.cur, tab = S.tab) => !!(M(ym).lock && M(ym).lock[tab]);
  const lockBar = () => locked() ? `<div class="lockbar" role="note"><b>저장 완료 — ${ymLong(S.cur)} ${TAB_NAME[S.tab]}은(는) 수정할 수 없습니다.</b>
    <span>${S.cur < FIRST_PAY_MONTH ? '이전 관리 파일에서 옮긴 기록입니다.' : ''}</span><span>고치려면 위쪽의 「수정」 버튼을 누르세요.</span></div>` : '';

  /* ---------- 화면 그리기 · 입력 처리 ---------- */
  const app = document.getElementById('app');
  function render() {
    month(S.cur);
    document.getElementById('ym').textContent = ymLong(S.cur);
    document.getElementById('saved').textContent = S.savedAt ? `마지막 저장 ${S.savedAt}${S.fileName ? ' → ' + S.fileName : ''}` : '아직 엑셀에 저장하지 않았습니다';
    app.innerHTML = kpis() + `<div class="tabs" role="tablist">
      <button role="tab" aria-selected="${S.tab === 'meter'}" class="${S.tab === 'meter' ? 'on' : ''}" data-act="tab" data-v="meter">① 전기 검침 (매월 10일)</button>
      <button role="tab" aria-selected="${S.tab === 'pay'}" class="${S.tab === 'pay' ? 'on' : ''}" data-act="tab" data-v="pay">② 임대료·관리비 납부현황</button><a class="tablink" href="#policy" data-act="gopolicy">요금 계산 기준 보기 ↓</a></div>` + (S.hasData ? '' : `<div class="lockbar" role="note"><b>이전 기록이 아직 없습니다.</b><span>받은 parkville-data.js 파일을 한 번 불러오면 이름·임대료·지난 기록이 채워집니다.</span><button type="button" class="btn" data-act="import">기록 파일 불러오기</button></div>`) + lockBar() + `<div class="lockable">` + (S.tab === 'meter' ? viewMeter() : viewPay()) + `</div>`;
    if (locked()) app.querySelectorAll('.lockable input, .lockable .seg button').forEach((el) => { el.disabled = true; });
    document.getElementById('editbtn').disabled = !locked();
    renderPanel();
    renderPolicy();
    persist();
  }
  // 화면 맨 아래 「요금 계산 기준」 — 지금 적용 중인 금액으로 보여 준다
  function renderPolicy() {
    const m = M(S.cur), items = m.fixedItems, mo = +S.cur.slice(5), pm = +shift(S.cur, -1).slice(5);
    const f1 = byFloor(1).map((room) => `${room.label} ${won(calc(room).fixed) || 0}원`).join(' · ');
    const chk = '<em class="chk">확인 필요</em>';
    const box = document.getElementById('policy'); if (!box) return;
    box.innerHTML = `<h2>요금 계산 기준 <small>${ymLong(S.cur)} 적용 기준</small></h2>
      <div class="pol">
        <section><h3>1. 임대료</h3><ul>
          <li>호실마다 정해진 월 임대료를 받습니다. 금액은 납부현황 표에서 고칠 수 있고, 고친 달부터 이어집니다.</li>
          <li>${h(S.cfg.notice)}</li></ul></section>
        <section><h3>2. 관리비 = 고정비 + 사용료</h3><ul>
          <li><b>2~5층 고정비 ${won(fixedSum(S.cur))}원</b> = ${items.map((x) => `${x[0]} ${won(x[1])}원`).join(' + ')}</li>
          <li><b>1층은 일반관리비만 고정</b>이고 전기 기본요금·TV 수신료·인터넷 요금은 없습니다. (${f1})</li>
          <li>사용료는 아래 3번의 전기 사용료입니다.</li></ul></section>
        <section><h3>3. 전기 사용료</h3><ul>
          <li>매월 10일에 검침합니다. 사용기간은 전월 10일 ~ 당월 9일입니다. (이번 달: ${pm}/10 ~ ${mo}/09)</li>
          <li>호실별 사용량(kWh)을 넣습니다. 사용량 = 당월 검침 − 전월 검침.</li>
          <li>한전 전기요금은 <b>층마다 따로</b> 나옵니다. 한전에 전화해서 확인한 2·3·4·5층 요금을 각각 넣습니다.</li>
          <li><b>호실 사용 요금 = 그 층 한전 요금 × (호실 사용량 ÷ 그 층 사용량)</b></li>
          <li><b>공용 전기요금은 건물 전체 1건</b>입니다. 2~5층 ${METERED.length}개 호실에 똑같이 나눠 공동전기료로 받습니다.</li>
          <li><b>1층은 상가</b>라서 전기요금을 입주자가 직접 냅니다. 검침·사용료를 입력하지 않습니다.</li></ul></section>
        <section><h3>4. 완납 / 미납과 가산료</h3><ul>
          <li>임대료와 관리비를 각각 완납 또는 미납으로 표시합니다.</li>
          <li><b>납부 마감일(매월 ${h(S.cfg.dueDay)}일)까지 내지 않으면 연체료 ${LATE_RATE * 100}%</b>가 붙습니다. 미납으로 표시하면 다음 달 「전월 미납 + 가산료」 칸에 나타납니다.</li>
          <li>전화해서 <b>기간 안에 받으면 완납으로 바꾸고 가산료는 없습니다</b> (「기간 내 받음」). 가산료까지 받았으면 「가산료 포함 받음」을 누르면 비고에 기록됩니다.</li>
          <li>임대료 미납과 관리비 미납 모두에 붙습니다.</li></ul></section>
        <section><h3>5. 납부영수증</h3><ul>
          <li>일반관리비 소계, 전기요금(기본 요금·사용량·사용 요금·공동전기료·TV수신료) 소계, 통신요금(인터넷.유선요금), 당월 요금 합계 순서로 찍힙니다.</li>
          <li>미납금액 = 이전 달 미납 관리비 + 가산료 ${LATE_RATE * 100}% · 납기내금액 = 당월 요금 합계 + 미납금액</li>
          <li>연체료(${LATE_RATE * 100}%) = 당월 요금 합계의 ${LATE_RATE * 100}% · 납기후금액 = 납기내금액 + 연체료</li>
          <li>납부 마감일은 매월 ${h(S.cfg.dueDay)}일입니다. 영수증은 엑셀에서 A4 가로 한 장에 3장씩 인쇄됩니다.</li></ul></section>
      </div>
`;
  }
  // 오른쪽 영수증 미리보기
  const panel = document.getElementById('panel');
  function renderPanel() {
    app.querySelectorAll('tr[data-row]').forEach((tr) => tr.classList.toggle('sel', tr.dataset.row === S.sel));
    const room = S.rooms.find((x) => x.id === S.sel);
    if (!room) { panel.innerHTML = `<div class="rc-empty">호실 줄을 누르면 이곳에 납부영수증 미리보기가 나옵니다.</div>`; return; }
    const r = receipt(room), n = (v) => (v ? won(v) : ''), row = (a, b, cls = '') => `<tr class="${cls}"><td>${a}</td><td class="n">${b}</td></tr>`;
    panel.innerHTML = `<div class="rc-head"><b>${room.label} 납부영수증 미리보기</b><span>${ymShort(S.cur)}</span></div>
      <table class="rc">
        <tr><th colspan="3">납부영수증(입주자용)</th></tr>
        <tr><td class="c">${room.label}</td><td class="c">${h(r.name) || '&nbsp;'} 님</td><td class="c">파크빌</td></tr>
        <tr><td colspan="3" class="c">사용기간(${r.period})</td></tr>
        <tr class="g"><td class="c">${r.y} 년</td><td class="c">${String(r.mo).padStart(2, '0')} 월</td><td class="c">관리비</td></tr>
        <tr><td colspan="2" class="c">구 분</td><td class="c">금 액</td></tr>
        <tr><td rowspan="4" class="c">일 반<br>관 리 비</td><td class="c">전기안전관리비</td><td></td></tr>
        <tr><td class="c">소방안전관리비</td><td></td></tr><tr><td class="c">L/V 안전관리비</td><td></td></tr><tr><td class="c">수 도 요 금</td><td></td></tr>
        <tr><td class="g"></td><td class="c g">소 계</td><td class="n">${n(r.general)}</td></tr>
        <tr><td rowspan="7" class="c">전기요금</td><td class="c">기본 요금</td><td class="n">${n(r.base)}</td></tr>
        <tr><td class="c">사 용 량</td><td class="n">${r.usage != null ? won(r.usage) : ''} kw</td></tr>
        <tr><td class="c">사용 요금</td><td class="n">${n(r.use)}</td></tr><tr><td class="c">공동전기료</td><td class="n">${n(r.common)}</td></tr>
        <tr><td class="c">전기부과세</td><td></td></tr><tr><td class="c">기금,기타</td><td></td></tr><tr><td class="c">TV수신료</td><td class="n">${n(r.tv)}</td></tr>
        <tr><td class="g"></td><td class="c g">소 계</td><td class="n">${n(r.elecSub)}</td></tr>
        <tr><td rowspan="2" class="c">통신요금</td><td class="c">인터넷.유선요금</td><td class="n">${n(r.net)}</td></tr><tr><td>&nbsp;</td><td></td></tr>
        <tr><td colspan="2" class="c g">당월 요금 합계</td><td class="n strong">${n(r.total)}</td></tr>
        <tr><td colspan="2" class="c g">미 납 금 액(${r.unpaidLabel})</td><td class="n">${n(r.unpaid)}</td></tr>
        <tr><td colspan="2" class="c">납 기 내 금 액</td><td class="n strong">${n(r.within)}</td></tr>
        <tr><td colspan="2" class="c">연 체 료 (5%)</td><td class="n">${n(r.late)}</td></tr>
        <tr><td colspan="2" class="c g">납 기 후 금 액</td><td class="n">${n(r.after)}</td></tr>
        <tr><td colspan="3" class="c g">${h(S.cfg.notice)}</td></tr>
        <tr><td colspan="3" class="c">입금계좌 : ${h(S.cfg.bank)}</td></tr><tr><td colspan="3" class="c">예금주 : ${h(S.cfg.holder)}</td></tr>
        <tr><td colspan="3" class="c">납부 마감일 ${r.due}</td></tr><tr><td colspan="3" class="c">파크빌 관리사무실 &nbsp; 수납인</td></tr>
      </table>
      <p class="hint">저장하면 엑셀의 「${ymShort(S.cur)} 영수증」 시트에 전 호실 영수증이 A4 가로 한 장에 3장씩 들어갑니다.</p>
      <details class="rc-cfg"><summary>영수증 문구 고치기 (계좌·예금주·마감일)</summary>
        <label>입금계좌 <input class="in note" data-cfg="bank" value="${h(S.cfg.bank)}"></label><label>예금주 <input class="in note" data-cfg="holder" value="${h(S.cfg.holder)}"></label>
        <label>납부 마감일(매월) <input class="in" data-cfg="dueDay" inputmode="numeric" value="${h(S.cfg.dueDay)}"></label><label>안내 문구 <input class="in note" data-cfg="notice" value="${h(S.cfg.notice)}"></label></details>`;
    if (cfgOpen) panel.querySelector('.rc-cfg').open = true;
  }
  let cfgOpen = false;
  const pickRow = (e) => { const tr = e.target.closest && e.target.closest('tr[data-row]'); if (tr && tr.dataset.row !== S.sel) { S.sel = tr.dataset.row; renderPanel(); persist(); } };
  app.addEventListener('click', pickRow); app.addEventListener('focusin', pickRow);
  let pendingFocus = null;
  function rerender() {
    setTimeout(() => {
      const a = document.activeElement, k = pendingFocus || (a && a.dataset && a.dataset.k), y = window.scrollY;
      pendingFocus = null; render(); window.scrollTo(0, y);
      if (k) { const el = document.querySelector(`[data-k="${k}"]`); if (el) { el.focus(); if (el.select) el.select(); } }
    }, 0);
  }
  function toast(msg, bad) {
    const t = document.getElementById('toast'); t.textContent = msg; t.className = 'show' + (bad ? ' bad' : '');
    clearTimeout(toast.t); toast.t = setTimeout(() => (t.className = ''), 5000);
  }
  // 자료 파일(parkville-data.js)을 읽어 이름·임대료·지난 기록을 채운다. 파일은 이 컴퓨터 안에서만 읽힌다.
  function importData(el) {
    const file = el.files && el.files[0]; if (!file) return;
    const rd = new FileReader();
    rd.onload = () => {
      el.value = '';
      let data;
      try { const t = String(rd.result); data = JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1)); } catch (err) { data = null; }
      if (!data || !Array.isArray(data.rooms) || !data.months) return toast('파크빌 기록 파일이 아닙니다. parkville-data.js 파일을 골라 주세요.', true);
      if (S.hasData && !confirm('지금까지 입력한 내용을 지우고 기록 파일 내용으로 바꿉니다. 계속할까요?')) return;
      S = seed(data); render(); toast('이전 기록을 불러왔습니다.');
    };
    rd.readAsText(file, 'utf-8');
  }
  function settle(room, how) {
    const a = arrears(room); if (!a.principal) return;
    a.items.forEach((x) => { rec(x.ym, room.id)[x.f] = 'Y'; });
    if (how === 'late') {
      const r = rec(S.cur, room.id), line = `전월 미납 ${won(a.principal)} + 가산료 ${won(a.late)} 수납`;
      r.note = r.note ? r.note + ' / ' + line : line;
    }
    toast(`${room.label} 전월 미납분을 완납 처리했습니다${how === 'late' ? ' (가산료 수납은 비고에 기록)' : ' (가산료 없음)'}.`);
  }

  document.addEventListener('click', (e) => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const { act, v, room, f } = b.dataset;
    if (b.tagName === 'A') e.preventDefault();
    if (act === 'pay') { if (locked()) return; const r = rec(S.cur, room); r[f] = r[f] === v ? '' : v; }
    else if (act === 'unlock') { if (!locked()) return; month(S.cur).lock[S.tab] = false; toast(`${ymLong(S.cur)} ${TAB_NAME[S.tab]}을(를) 다시 수정할 수 있습니다. 고친 뒤 「저장」을 눌러 주세요.`); }
    else if (act === 'tab') S.tab = v;
    else if (act === 'month') S.cur = shift(S.cur, +v);
    else if (act === 'settle') settle(S.rooms.find((x) => x.id === room), v);
    else if (act === 'undo') return undo();
    else if (act === 'gopolicy') return document.getElementById('policy').scrollIntoView({ behavior: 'smooth' });
    else if (act === 'import') return document.getElementById('datafile').click();
    else if (act === 'save') return saveExcel(false);
    else if (act === 'pickfile') return saveExcel(true);
    else if (act === 'reset') { if (confirm('입력한 내용을 모두 지우고 처음 상태로 되돌릴까요?')) S = seed(); else return; }
    render();
  });
  document.addEventListener('change', (e) => {
    const el = e.target, f = el.dataset.f;
    if (el.id === 'datafile') return importData(el);
    if (el.dataset.cfg) { S.cfg[el.dataset.cfg] = el.value.trim(); cfgOpen = true; persist(); return renderPanel(); }
    if (!f || locked()) return;
    if (f === 'common') { month(S.cur).common = num(el.value); return rerender(); }
    if (f === 'kepcoF') { const m = month(S.cur); (m.kepcoF || (m.kepcoF = {}))[el.dataset.floor] = num(el.value); return rerender(); }
    if (f === 'fixedItem') { month(S.cur).fixedItems[+el.dataset.i][1] = num(el.value) || 0; fixedOpen = true; return rerender(); }
    const r = rec(S.cur, el.dataset.room);
    if (f === 'note') r.note = el.value.trim();
    else if (f === 'tenant') r.tenant = el.value.trim();
    else if (f === 'read') { r.read = num(el.value); delete r.usage; }
    else if (f === 'usage') {
      const n = num(el.value), prev = peek(shift(S.cur, -1), el.dataset.room).read;
      if (n == null) delete r.usage; else { r.usage = n; if (prev != null) r.read = prev + n; }
    }
    else if (f === 'fee') {
      const n = num(el.value);
      // 이전 파일에서 옮긴 달은 관리비 총액으로 들어 있으므로, 사용료를 고치면 총액을 함께 고친다
      if (r.mgmtAmt != null) r.mgmtAmt = calc(S.rooms.find((x) => x.id === el.dataset.room)).fixed + (n || 0);
      else if (n == null) delete r.fee; else r.fee = n;
    }
    else if (f === 'rent' || f === 'fixed') { const n = num(el.value); if (n != null) { r[f] = n; if (f === 'fixed') delete r.mgmtAmt; } }
    rerender();
  });
  let fixedOpen = false; // 고정비 내역을 고치는 동안에는 접히지 않게
  new MutationObserver(() => { const d = app.querySelector('.fixedbox'); if (d && fixedOpen) d.open = true; }).observe(app, { childList: true });
  document.addEventListener('toggle', (e) => { if (e.target.classList && e.target.classList.contains('fixedbox')) fixedOpen = e.target.open; }, true);
  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter' || !e.target.classList.contains('in') || !e.target.dataset.room) return;
    const f = e.target.dataset.f, list = [...app.querySelectorAll(`input[data-f="${f}"]`)], i = list.indexOf(e.target);
    if (list[i + 1]) pendingFocus = list[i + 1].dataset.k;
    e.target.blur(); if (pendingFocus) rerender();
  });

  /* ---------- 엑셀 저장 ---------- */
  const C = (v, s) => ({ v, s });
  function paySheet(ym) {
    const H = (x) => C(x, ST.head), m = M(ym);
    const rows = [[C(`파크빌 ${ymLong(ym)} 임대료·관리비 납부현황`, ST.title)], [C(`작성일: ${new Date().toLocaleDateString('ko-KR')}`, ST.sub)],
      ['호수', '이름', '임대료', '완납/미납', '관리비', '', '', '완납/미납', '전월 미납', '가산료(5%)', '비고'].map(H),
      ['', '', '', '', '고정', '사용료', '계', '', '', '', ''].map(H)];
    const pc = (v) => C(paidText[v], v === 'N' ? ST.bad : v === 'Y' ? ST.good : ST.center), t = summary(ym);
    S.rooms.forEach((room) => {
      const c = calc(room, ym), a = arrears(room, ym);
      rows.push([C(room.label, ST.center), C(c.tenant, ST.left), C(c.rent, ST.num), pc(c.rentPaid), C(c.fixed, ST.num), C(c.fee == null ? '' : c.fee, ST.num), C(c.mgmt, ST.num), pc(c.mgmtPaid),
        C(a.principal || '', ST.num), C(a.late || '', ST.num), C(c.note, ST.left)]);
    });
    rows.push([C('합계', ST.sumLabel), C('', ST.sumLabel), C(t.rentSum, ST.sumNum), C(`완납 ${t.rentY} / 미납 ${t.rentN}`, ST.sumLabel), C(t.fixedSum, ST.sumNum), C(t.feeSum, ST.sumNum), C(t.mgmtSum, ST.sumNum),
      C(`완납 ${t.mgmtY} / 미납 ${t.mgmtN}`, ST.sumLabel), C(t.arrP || '', ST.sumNum), C(t.arrL || '', ST.sumNum), C('', ST.sumLabel)]);
    const n = rows.length;
    rows.push([], [C(`※ 2~5층 고정비 ${won(fixedSum(ym))}원 = ${m.fixedItems.map((x) => x[0] + ' ' + won(x[1])).join(' + ')}   /   1층은 일반관리비만 고정`, ST.sub)],
      [C(`※ 사용료 = 층별 한전 전기요금 × (호실 사용량 ÷ 그 층 사용량) + 공동전기료(공용 요금 ÷ ${METERED.length}호실)   /   마감일까지 미납 시 연체료 5%`, ST.sub)]);
    return { name: `${ymShort(ym)} 납부현황`, widths: [11, 16, 12, 12, 11, 11, 11, 12, 12, 11, 34], heights: { 0: 30 }, freeze: 4, rows,
      merges: ['A1:K1', 'A2:K2', 'A3:A4', 'B3:B4', 'C3:C4', 'D3:D4', 'E3:G3', 'H3:H4', 'I3:I4', 'J3:J4', 'K3:K4', `A${n + 2}:K${n + 2}`, `A${n + 3}:K${n + 3}`] };
  }
  function meterSheet(ym) {
    const [y, mo] = ym.split('-').map(Number), p = shift(ym, -1), m = M(ym), tot = totalUsage(ym);
    const rows = [[C(`파크빌 ${ymLong(ym)} 전기검침`, ST.title)],
      [C(`검침일: ${y}.${mo}.10   사용기간: ${+p.slice(5)}/10 ~ ${mo}/09   한전 요금: ${[2, 3, 4, 5].map((f) => `${f}층 ${floorKepco(ym, f) != null ? won(floorKepco(ym, f)) : '미입력'}`).join(' / ')} / 공용 ${m.common != null ? won(m.common) : '미입력'}   전체 사용량: ${won(tot) || 0}kWh`, ST.sub)],
      ['호수', '이름', '전월 검침', '당월 검침', '사용량(kWh)', '전기 사용료(원)'].map((x) => C(x, ST.head))];
    FLOORS.slice(1).forEach((f) => {
      let u = 0, a = 0;
      byFloor(f).forEach((room) => {
        const c = calc(room, ym); u += c.usage > 0 ? c.usage : 0; a += c.autoFee || 0;
        rows.push([C(room.label, ST.center), C(c.tenant, ST.left), C(c.prev ?? '', ST.num), C(c.read ?? '', ST.num), C(c.usage ?? '', ST.num), C(c.autoFee ?? '', ST.num)]);
      });
      rows.push([C(`${f}층 소계`, ST.sumLabel), C(floorKepco(ym, f) != null ? `한전 요금 ${won(floorKepco(ym, f))}원` : '', ST.sumLabel), C('', ST.sumLabel), C('', ST.sumLabel), C(u, ST.sumNum), C(a || '', ST.sumNum)]);
    });
    return { name: `${ymShort(ym)} 전기검침`, widths: [10, 18, 13, 13, 13, 16], heights: { 0: 30 }, freeze: 3, rows, merges: ['A1:F1', 'A2:F2'] };
  }
  function historySheet() {
    const yms = Object.keys(S.months).filter((ym) => METERED.some((id) => peek(ym, id).read != null)).sort();
    const last = colName(1 + yms.length * 2), cur = yms[yms.length - 1] || S.cur;
    const r3 = [C('호수', ST.head), C('이름', ST.head)], r4 = [C('', ST.head), C('', ST.head)], merges = [`A1:${last}1`, `A2:${last}2`, 'A3:A4', 'B3:B4'];
    yms.forEach((ym, i) => { r3.push(C(ymShort(ym), ST.head), C('', ST.head)); r4.push(C('검침', ST.head), C('사용량', ST.head)); merges.push(`${colName(2 + i * 2)}3:${colName(3 + i * 2)}3`); });
    const rows = [[C('파크빌 호수별 전기검침 기록', ST.title)], [C('매월 10일 검침 · 사용량 = 당월 검침 − 전월 검침', ST.sub)], r3, r4];
    S.rooms.filter((r) => r.floor > 1).forEach((room) => {
      const row = [C(room.label, ST.center), C(calc(room, cur).tenant, ST.left)];
      yms.forEach((ym) => { const c = calc(room, ym); row.push(C(c.read ?? '', ST.num), C(c.usage ?? '', ST.num)); });
      rows.push(row);
    });
    return { name: '호수별 전기검침 기록', widths: [10, 18, ...yms.flatMap(() => [11, 9])], heights: { 0: 30 }, freeze: 4, rows, merges };
  }
  function receiptSheet(ym) {
    const N = 30, W = 11, rows = [], merges = [], breaks = [];
    S.rooms.forEach((room, i) => {
      const blk = Math.floor(i / 3), k = i % 3, r0 = blk * N, c0 = k * 4, r = receipt(room, ym);
      if (k === 0) { for (let j = 0; j < N; j++) rows.push(new Array(W).fill(null)); if (blk) breaks.push(r0); }
      const col = (j) => colName(c0 + j), n = (v) => (v ? v : '');
      const put = (j, a, b, c, sa = ST.center, sb = ST.center, sc = ST.num) => { rows[r0 + j][c0] = C(a, sa); rows[r0 + j][c0 + 1] = C(b, sb); rows[r0 + j][c0 + 2] = C(c, sc); };
      const wide = (j, text, s) => { put(j, text, '', '', s, s, s); merges.push(`${col(0)}${r0 + j + 1}:${col(2)}${r0 + j + 1}`); };
      const two = (j, text, v, s, sv) => { put(j, text, '', v, s, s, sv); merges.push(`${col(0)}${r0 + j + 1}:${col(1)}${r0 + j + 1}`); };
      const tall = (j, len) => merges.push(`${col(0)}${r0 + j + 1}:${col(0)}${r0 + j + len}`);
      wide(0, '납부영수증(입주자용)', ST.boldC);
      put(1, room.label, (r.name || '      ') + ' 님', '파크빌', ST.center, ST.center, ST.center);
      wide(2, `사용기간(${r.period})`, ST.center);
      put(3, `${r.y} 년`, `${String(r.mo).padStart(2, '0')} 월`, '관리비', ST.gray, ST.gray, ST.gray);
      two(4, '구      분', '금    액', ST.center, ST.center);
      put(5, '일 반 관리비', '전기안전관리비', ''); put(6, '', '소방안전관리비', ''); put(7, '', 'L/V 안전관리비', ''); put(8, '', '수 도 요 금', ''); tall(5, 4);
      put(9, '', '소      계', n(r.general), ST.gray, ST.gray);
      put(10, '전기요금', '기본 요금', n(r.base)); put(11, '', '사 용 량', (r.usage != null ? won(r.usage) : '') + ' kw'); put(12, '', '사용 요금', n(r.use)); put(13, '', '공동전기료', n(r.common));
      put(14, '', '전기부과세', ''); put(15, '', '기금,기타', ''); put(16, '', 'TV수신료', n(r.tv)); tall(10, 7);
      put(17, '', '소      계', n(r.elecSub), ST.gray, ST.gray);
      put(18, '통신요금', '인터넷.유선요금', n(r.net)); put(19, '', '', ''); tall(18, 2);
      two(20, '당월 요금 합계', n(r.total), ST.gray, ST.grayNum);
      two(21, `미 납 금 액(${r.unpaidLabel})`, n(r.unpaid), ST.gray, ST.num);
      two(22, '납 기 내 금 액', n(r.within), ST.boldC, ST.num);
      two(23, '연 체 료 (5%)', n(r.late), ST.center, ST.num);
      two(24, '납 기 후 금 액', n(r.after), ST.gray, ST.num);
      wide(25, S.cfg.notice, ST.gray); wide(26, `입금계좌 : ${S.cfg.bank}`, ST.center); wide(27, `예금주 : ${S.cfg.holder}`, ST.center);
      wide(28, `납부 마감일 ${r.due}`, ST.center); wide(29, '파크빌 관리사무실      수납인', ST.center);
    });
    // A4 가로 한 장에 영수증 3장
    return { name: `${ymShort(ym)} 영수증`, widths: [12, 19, 14, 3, 12, 19, 14, 3, 12, 19, 14], rowHeight: 17.5, fitWidth: true, landscape: true, rows, merges, breaks };
  }
  function monthlySheet(yms) {
    const rows = [[C('파크빌 월별 요약', ST.title)], [C('임대료·관리비는 그 달 표에 적힌 금액의 합계입니다.', ST.sub)],
      ['월', '임대료 합계', '임대료 완납(호)', '관리비 합계', '관리비 완납(호)', '한전 전기요금'].map((x) => C(x, ST.head))];
    yms.forEach((ym) => { const t = summary(ym); rows.push([C(ymShort(ym), ST.center), C(t.rentSum, ST.num), C(t.rentY, ST.num), C(t.mgmtSum, ST.num), C(t.mgmtY, ST.num), C(kepcoSum(ym) ?? '', ST.num)]); });
    return { name: '월별 요약', widths: [11, 15, 15, 15, 15, 16], heights: { 0: 30 }, freeze: 3, rows, merges: ['A1:F1', 'A2:F2'] };
  }
  function buildWorkbook() {
    const has = (ym, f) => S.rooms.some((room) => peek(ym, room.id)[f] != null);
    const yms = Object.keys(S.months).filter((ym) => ym === S.cur || ['rentPaid', 'mgmtPaid', 'mgmtAmt', 'tenant', 'fee', 'note'].some((f) => has(ym, f)) || (ym >= FIRST_PAY_MONTH && has(ym, 'read'))).sort().reverse();
    const sheets = [];
    yms.forEach((ym) => { sheets.push(paySheet(ym)); if (ym === S.cur || has(ym, 'read') || has(ym, 'usage')) sheets.push(meterSheet(ym)); });
    return workbook([...sheets.slice(0, 2), receiptSheet(S.cur), monthlySheet(yms), historySheet(), ...sheets.slice(2)]);
  }

  // 한 번 고른 파일을 기억해 두었다가 다음부터는 버튼만 누르면 같은 파일에 덮어쓴다
  const idb = (mode, fn) => new Promise((res, rej) => {
    const o = indexedDB.open('parkville', 1);
    o.onupgradeneeded = () => o.result.createObjectStore('kv');
    o.onerror = () => rej(o.error);
    o.onsuccess = () => { const tx = o.result.transaction('kv', mode), rq = fn(tx.objectStore('kv')); tx.oncomplete = () => res(rq.result); tx.onerror = () => rej(tx.error); };
  });
  async function saveExcel(forcePick) {
    const willLock = !forcePick && !locked();
    const bytes = buildWorkbook();
    const stamp = () => { if (willLock) { const m = month(S.cur); (m.lock || (m.lock = {}))[S.tab] = true; } S.savedAt = new Date().toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }); };
    if (!window.showSaveFilePicker) { // 사파리 등: 다운로드 폴더로 내려받기
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      a.download = FILE_NAME; a.click(); stamp(); render();
      return toast(`다운로드 폴더에 ${FILE_NAME} 로 저장했습니다. (크롬·엣지에서는 정해 둔 파일에 바로 저장됩니다)`);
    }
    let handle = null;
    try {
      if (!forcePick) handle = await idb('readonly', (s) => s.get('file')).catch(() => null);
      if (!handle) {
        handle = await window.showSaveFilePicker({ id: 'parkville', startIn: 'desktop', suggestedName: FILE_NAME, types: [{ description: '엑셀 파일', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }] });
        await idb('readwrite', (s) => s.put(handle, 'file')).catch(() => { });
      }
      if ((await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') return toast('파일 저장 권한을 허용해야 저장할 수 있습니다.', true);
      const w = await handle.createWritable(); await w.write(bytes); await w.close();
      // 정말 써졌는지 파일을 다시 읽어 크기를 확인한다
      const written = await handle.getFile();
      if (written.size !== bytes.length) throw new Error('size mismatch');
      S.fileName = handle.name; stamp(); render();
      toast(`${handle.name} 파일에 저장했습니다. (${new Date(written.lastModified).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' })})`);
    } catch (err) {
      if (err.name === 'AbortError') return;
      if (err.name === 'NotFoundError' && !forcePick) { await idb('readwrite', (s) => s.delete('file')).catch(() => { }); return saveExcel(true); }
      toast('저장하지 못했습니다.', true);
      alert('엑셀 파일에 저장하지 못했습니다.\n\n엑셀에서 그 파일을 열어 두었다면 먼저 닫고, 다시 「저장」을 눌러 주세요.\n(화면에 입력한 내용은 그대로 남아 있습니다)');
    }
  }

  // 자동 점검용(#test…): 저장소를 건드리지 않고 예시 값을 넣어 화면과 엑셀 내용을 확인한다
  if (TEST) {
    const mode = location.hash.slice(6);
    if (mode) month('2026-11').kepcoF = { 2: 255460, 3: 282400, 4: 260680, 5: 246480 };
    if (mode === 'meter') S.tab = 'meter';
    if (mode === 'hist') S.cur = '2025-03';
    if (mode === 'undo') { render(); const a = peek('2026-11', '201').read; rec('2026-11', '201').read = 99999; rec('2026-11', '202').rentPaid = 'N'; render(); const b = peek('2026-11', '201').read; S.cur = '2026-12'; render(); undo(); document.body.dataset.undo = [a, b, peek('2026-11', '201').read, peek('2026-11', '202').rentPaid || '-', S.cur, undoStack.length].join('|'); }
    if (mode === 'rc') { month('2026-11').common = 132550; S.sel = '205'; }
    if (mode.startsWith('dec')) { S.cur = '2026-12'; S.tab = 'meter'; Object.assign(rec('2026-12', '201'), { usage: 150, read: 50378 }); rec('2026-12', '202').read = 38400; month('2026-12').lock = { meter: mode === 'decL' }; }
    if (mode === 'late' || mode === 'settle') {
      Object.assign(rec('2026-11', '201'), { rentPaid: 'N', mgmtPaid: 'N' }); rec('2026-11', '305').mgmtPaid = 'N'; rec('2026-11', '202').rentPaid = 'Y';
      rec('2026-11', '204').tenant = '홍길동'; S.cur = '2026-12';
    }
    render();
    if (mode === 'settle') { settle(S.rooms.find((x) => x.id === '201'), 'late'); settle(S.rooms.find((x) => x.id === '305'), 'in'); render(); }
    const b = buildWorkbook(); let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
    const pre = document.createElement('pre'); pre.id = 'testout'; pre.hidden = true; pre.textContent = btoa(s); document.body.appendChild(pre);
  } else render();
})();
