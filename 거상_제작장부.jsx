import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';

/* ────────────────────────────────────────────────────────────
   숫자 다루기
   거상은 억 단위가 기본이라, 입력은 관대하게 받고 출력은 읽기 쉽게.
   "16억", "1억 5000만", "100,250,000" 모두 같은 방식으로 해석한다.
   ──────────────────────────────────────────────────────────── */

function parseAmount(raw) {
  if (typeof raw === 'number') return Math.round(raw);
  let t = String(raw ?? '').replace(/[\s,원]/g, '');
  if (t === '') return 0;
  if (/[억만천]/.test(t)) {
    let total = 0;
    const re = /(\d*\.?\d+)\s*(억|만|천)?/g;
    let m;
    while ((m = re.exec(t))) {
      const v = parseFloat(m[1]);
      if (isNaN(v)) continue;
      const u = m[2];
      total += u === '억' ? v * 1e8 : u === '만' ? v * 1e4 : u === '천' ? v * 1e3 : v;
    }
    return Math.round(total);
  }
  const n = parseFloat(t);
  return isNaN(n) ? 0 : Math.round(n);
}

const comma = (n) => (n || 0).toLocaleString('ko-KR');

function korean(n) {
  const neg = n < 0;
  let v = Math.abs(Math.round(n || 0));
  if (v === 0) return '0';
  const eok = Math.floor(v / 1e8);
  const man = Math.floor((v % 1e8) / 1e4);
  const rest = v % 1e4;
  const parts = [];
  if (eok) parts.push(`${comma(eok)}억`);
  if (man) parts.push(`${comma(man)}만`);
  if (rest) parts.push(comma(rest));
  return (neg ? '-' : '') + parts.join(' ');
}

const uid = () => Math.random().toString(36).slice(2, 9);

/* ────────────────────────────────────────────────────────────
   처음 장부를 열었을 때 들어있는 예시 한 건
   ──────────────────────────────────────────────────────────── */

const SEED = {
  prices: { 봉인의돌: 100250000, 사원서판: 18000000, 힘의기억: 6000000 },
  items: [
    {
      id: uid(),
      name: '청혼상류봉',
      path: '장과로',
      fee: 200000000,
      count: 1,
      sellPrice: 2000000000,
      memo: '',
      open: true,
      mats: [
        { id: uid(), name: '봉인의돌', qty: 2 },
        { id: uid(), name: '사원서판', qty: 20 },
        { id: uid(), name: '힘의기억', qty: 150 },
      ],
    },
  ],
};

const STORAGE_KEY = 'gersang:ledger:v1';
const SYNC_KEY = 'gersang:sync:v1';
const GIST_FILE = 'gersang-ledger.json';

/* 보관함 — 이 자리는 웹 배포용으로 빌드할 때 localStorage 로 바뀐다 */
const store = {
  get: (k) => window.storage.get(k),
  set: (k, v) => window.storage.set(k, v),
};

/* ────────────────────────────────────────────────────────────
   GitHub Gist — 비공개 Gist 한 칸을 장부 보관함으로 쓴다
   ──────────────────────────────────────────────────────────── */

async function gistCall(token, path, opts = {}) {
  const r = await fetch('https://api.github.com' + path, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json',
    },
  });
  if (!r.ok) {
    const msg =
      r.status === 401
        ? '토큰이 맞지 않거나 만료됐습니다'
        : r.status === 403
        ? '요청이 너무 잦습니다. 잠시 뒤 다시 시도하세요'
        : r.status === 404
        ? 'Gist를 찾을 수 없습니다'
        : `연결 실패 (${r.status})`;
    throw new Error(msg);
  }
  return r.json();
}

async function gistRead(token, id) {
  const j = await gistCall(token, `/gists/${id}`);
  const f = j.files?.[GIST_FILE];
  if (!f) throw new Error('이 Gist 안에 장부 파일이 없습니다');
  const text = f.truncated ? await (await fetch(f.raw_url)).text() : f.content;
  return JSON.parse(text);
}

const gistWrite = (token, id, payload) =>
  gistCall(token, `/gists/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({ files: { [GIST_FILE]: { content: JSON.stringify(payload) } } }),
  });

const gistCreate = (token, payload) =>
  gistCall(token, '/gists', {
    method: 'POST',
    body: JSON.stringify({
      description: '거상 제작 장부',
      public: false,
      files: { [GIST_FILE]: { content: JSON.stringify(payload) } },
    }),
  });

/* ────────────────────────────────────────────────────────────
   금액 입력칸 — 볼 때는 콤마, 고칠 때는 날것
   ──────────────────────────────────────────────────────────── */

function AmountInput({ value, onCommit, placeholder, cls = '', label }) {
  const [focused, setFocused] = useState(false);
  const [draft, setDraft] = useState('');
  const shown = focused ? draft : value ? comma(value) : '';
  return (
    <input
      className={`amt ${cls}`}
      value={shown}
      aria-label={label}
      placeholder={placeholder}
      inputMode="decimal"
      onFocus={() => {
        setFocused(true);
        setDraft(value ? String(value) : '');
      }}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        setFocused(false);
        onCommit(parseAmount(draft));
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') e.currentTarget.blur();
      }}
    />
  );
}

/* ────────────────────────────────────────────────────────────
   본체
   ──────────────────────────────────────────────────────────── */

export default function App() {
  const [data, setData] = useState(null);
  const [tab, setTab] = useState('items');
  const [q, setQ] = useState('');
  const [status, setStatus] = useState('');
  const [confirmReset, setConfirmReset] = useState(false);
  const fileRef = useRef(null);
  const firstSave = useRef(true);

  const [sync, setSync] = useState({ token: '', gistId: '' });
  const [syncMsg, setSyncMsg] = useState('');
  const [busy, setBusy] = useState(false);
  const [tokenDraft, setTokenDraft] = useState('');
  const [gistDraft, setGistDraft] = useState('');
  const [dragId, setDragId] = useState(null);
  const [overId, setOverId] = useState(null);
  const stamp = useRef(0);
  const skipPush = useRef(false);
  const linked = !!(sync.token && sync.gistId);

  /* 불러오기 */
  useEffect(() => {
    let alive = true;
    (async () => {
      let local = SEED;
      try {
        const r = await store.get(STORAGE_KEY);
        const parsed = JSON.parse(r.value);
        if (parsed && parsed.items) local = parsed;
      } catch {}
      stamp.current = local.updatedAt || 0;
      if (!alive) return;
      setData(local);

      let cfg = { token: '', gistId: '' };
      try {
        const r = await store.get(SYNC_KEY);
        cfg = JSON.parse(r.value);
      } catch {}
      if (!alive || !cfg.token || !cfg.gistId) return;
      setSync(cfg);

      try {
        const remote = await gistRead(cfg.token, cfg.gistId);
        if (!alive) return;
        if ((remote.updatedAt || 0) > stamp.current) {
          stamp.current = remote.updatedAt || 0;
          skipPush.current = true;
          setData(remote);
          setSyncMsg('최신 장부를 가져왔습니다');
        } else {
          setSyncMsg('최신 상태입니다');
        }
      } catch (e) {
        if (alive) setSyncMsg(e.message);
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  /* 저장 — 잠깐 멈추면 알아서 */
  useEffect(() => {
    if (!data) return;
    if (firstSave.current) {
      firstSave.current = false;
      return;
    }
    if (skipPush.current) {
      skipPush.current = false;
      return;
    }
    const t = setTimeout(async () => {
      stamp.current = Date.now();
      const payload = { ...data, updatedAt: stamp.current };
      try {
        await store.set(STORAGE_KEY, JSON.stringify(payload));
        setStatus('저장됨');
      } catch {
        setStatus('저장 실패 — 이 기기에 보관되지 않습니다');
      }
      if (sync.token && sync.gistId) {
        try {
          await gistWrite(sync.token, sync.gistId, payload);
          setStatus('올렸습니다');
          setSyncMsg('');
        } catch (e) {
          setStatus('');
          setSyncMsg(e.message);
        }
      }
      setTimeout(() => setStatus(''), 1600);
    }, 900);
    return () => clearTimeout(t);
  }, [data, sync]);

  /* 동기화 조작 */
  const saveCfg = async (cfg) => {
    setSync(cfg);
    try {
      await store.set(SYNC_KEY, JSON.stringify(cfg));
    } catch {}
  };

  const connect = async (mode) => {
    const token = tokenDraft.trim();
    if (!token) return setSyncMsg('토큰을 붙여넣어 주세요');
    setBusy(true);
    setSyncMsg('');
    try {
      const payload = { ...data, updatedAt: Date.now() };
      if (mode === 'new') {
        const j = await gistCreate(token, payload);
        stamp.current = payload.updatedAt;
        await saveCfg({ token, gistId: j.id });
        setSyncMsg('새 보관함을 만들고 지금 장부를 올렸습니다');
      } else {
        const id = gistDraft.trim();
        if (!id) {
          setBusy(false);
          return setSyncMsg('Gist 주소나 아이디를 적어주세요');
        }
        const clean = id.split('/').filter(Boolean).pop();
        const remote = await gistRead(token, clean);
        stamp.current = remote.updatedAt || Date.now();
        skipPush.current = true;
        setData(remote);
        await saveCfg({ token, gistId: clean });
        setSyncMsg('기존 보관함의 장부를 불러왔습니다');
      }
      setTokenDraft('');
      setGistDraft('');
    } catch (e) {
      setSyncMsg(e.message);
    }
    setBusy(false);
  };

  const pullNow = async () => {
    setBusy(true);
    setSyncMsg('');
    try {
      const remote = await gistRead(sync.token, sync.gistId);
      stamp.current = remote.updatedAt || Date.now();
      skipPush.current = true;
      setData(remote);
      setSyncMsg('가져왔습니다');
    } catch (e) {
      setSyncMsg(e.message);
    }
    setBusy(false);
  };

  const pushNow = async () => {
    setBusy(true);
    setSyncMsg('');
    try {
      stamp.current = Date.now();
      await gistWrite(sync.token, sync.gistId, { ...data, updatedAt: stamp.current });
      setSyncMsg('올렸습니다');
    } catch (e) {
      setSyncMsg(e.message);
    }
    setBusy(false);
  };

  const unlink = async () => {
    await saveCfg({ token: '', gistId: '' });
    setSyncMsg('이 기기에서만 연결을 끊었습니다. 보관함과 장부는 그대로 있습니다');
  };

  const prices = data?.prices ?? {};
  const items = data?.items ?? [];

  /* 계산 — 재료 칸은 1개 기준, 수량만큼 곱한다 */
  const calc = (it) => {
    const n = Math.max(1, it.count || 1);
    const unitMat = it.mats.reduce((s, m) => s + (prices[m.name] || 0) * (m.qty || 0), 0);
    const unitCost = unitMat + (it.fee || 0);
    const profitUnit = it.sellPrice ? it.sellPrice - unitCost : null;
    return {
      n,
      unitMat,
      unitCost,
      mat: unitMat * n,
      fee: (it.fee || 0) * n,
      total: unitCost * n,
      profitUnit,
      profit: it.sellPrice ? profitUnit * n : null,
      margin: it.sellPrice ? profitUnit / it.sellPrice : null,
    };
  };

  const usage = useMemo(() => {
    const u = {};
    items.forEach((it) =>
      new Set(it.mats.map((m) => m.name).filter(Boolean)).forEach((n) => {
        u[n] = (u[n] || 0) + 1;
      })
    );
    return u;
  }, [items]);

  const grand = useMemo(
    () =>
      items.reduce(
        (a, it) => {
          const c = calc(it);
          a.cost += c.total;
          if (it.sellPrice) {
            a.sell += it.sellPrice * c.n;
            a.profit += c.profit;
          }
          return a;
        },
        { cost: 0, sell: 0, profit: 0 }
      ),
    [items, prices]
  );

  /* 검색 — 아이템 이름·제작 경로·메모, 그리고 재료 이름까지 훑는다 */
  const needle = q.trim().toLowerCase();

  const shownItems = useMemo(() => {
    if (!needle) return items.map((it) => ({ it, hitMats: [] }));
    return items
      .map((it) => {
        const hitMats = it.mats.map((m) => m.name).filter((n) => n && n.toLowerCase().includes(needle));
        const self = [it.name, it.path, it.memo].some((v) => (v || '').toLowerCase().includes(needle));
        return self || hitMats.length ? { it, hitMats } : null;
      })
      .filter(Boolean);
  }, [items, needle]);

  const shownMats = useMemo(() => {
    const all = Object.keys(prices).sort((a, b) => a.localeCompare(b, 'ko'));
    return needle ? all.filter((n) => n.toLowerCase().includes(needle)) : all;
  }, [prices, needle]);

  /* 고치기 */
  const patchItem = (id, patch) =>
    setData((d) => ({ ...d, items: d.items.map((it) => (it.id === id ? { ...it, ...patch } : it)) }));

  const patchMat = (itemId, matId, patch) =>
    setData((d) => ({
      ...d,
      items: d.items.map((it) =>
        it.id === itemId ? { ...it, mats: it.mats.map((m) => (m.id === matId ? { ...m, ...patch } : m)) } : it
      ),
    }));

  /* 여기가 핵심 — 가격은 재료 이름 하나당 한 곳에만 산다.
     어느 화면에서 고치든 그 재료를 쓰는 모든 아이템이 같이 움직인다. */
  const setPrice = (name, value) => {
    if (!name) return;
    setData((d) => ({ ...d, prices: { ...d.prices, [name]: value } }));
  };

  const addItem = () =>
    setData((d) => ({
      ...d,
      items: [
        ...d.items,
        { id: uid(), name: '', path: '', fee: 0, count: 1, sellPrice: 0, memo: '', open: true, mats: [{ id: uid(), name: '', qty: 1 }] },
      ],
    }));

  const dupItem = (it) =>
    setData((d) => ({
      ...d,
      items: [
        ...d.items,
        { ...it, id: uid(), name: it.name ? `${it.name} 사본` : '', open: true, mats: it.mats.map((m) => ({ ...m, id: uid() })) },
      ],
    }));

  const delItem = (id) => setData((d) => ({ ...d, items: d.items.filter((it) => it.id !== id) }));

  /* 순서 바꾸기 */
  const moveItem = (id, dir) =>
    setData((d) => {
      const i = d.items.findIndex((x) => x.id === id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= d.items.length) return d;
      const arr = [...d.items];
      [arr[i], arr[j]] = [arr[j], arr[i]];
      return { ...d, items: arr };
    });

  const dropOn = (fromId, toId) =>
    setData((d) => {
      if (!fromId || fromId === toId) return d;
      const arr = [...d.items];
      const from = arr.findIndex((x) => x.id === fromId);
      const to = arr.findIndex((x) => x.id === toId);
      if (from < 0 || to < 0) return d;
      const [moved] = arr.splice(from, 1);
      arr.splice(to, 0, moved);
      return { ...d, items: arr };
    });

  const addMat = (itemId) =>
    setData((d) => ({
      ...d,
      items: d.items.map((it) => (it.id === itemId ? { ...it, mats: [...it.mats, { id: uid(), name: '', qty: 1 }] } : it)),
    }));

  const delMat = (itemId, matId) =>
    setData((d) => ({
      ...d,
      items: d.items.map((it) => (it.id === itemId ? { ...it, mats: it.mats.filter((m) => m.id !== matId) } : it)),
    }));

  /* 내보내기 */
  const toExcel = () => {
    const wb = XLSX.utils.book_new();

    const s1 = [
      ['아이템 이름', '제작 경로', '제작 수량', '1개당 원가', '재료비', '수수료', '총 제작비', '판매가(1개)', '순이익', '마진율', '메모'],
    ];
    items.forEach((it) => {
      const c = calc(it);
      s1.push([
        it.name,
        it.path,
        c.n,
        c.unitCost,
        c.mat,
        c.fee,
        c.total,
        it.sellPrice || '',
        c.profit ?? '',
        c.margin ?? '',
        it.memo || '',
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s1), '아이템');

    const s2 = [['아이템명', '재료명', '개당 가격', '1개당 갯수', '제작 수량', '필요 갯수', '합계']];
    items.forEach((it) => {
      const c = calc(it);
      it.mats.forEach((m) => {
        const need = (m.qty || 0) * c.n;
        s2.push([it.name, m.name, prices[m.name] || 0, m.qty || 0, c.n, need, (prices[m.name] || 0) * need]);
      });
      if (it.fee) s2.push([it.name, '수수료', it.fee, 1, c.n, c.n, c.fee]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s2), '재료');

    const s3 = [['재료명', '개당 가격', '쓰는 아이템 수']];
    Object.keys(prices)
      .sort((a, b) => a.localeCompare(b, 'ko'))
      .forEach((n) => s3.push([n, prices[n], usage[n] || 0]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s3), '시세표');

    XLSX.writeFile(wb, '거상_제작장부.xlsx');
  };

  const backup = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '거상_제작장부_백업.json';
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const restore = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const fr = new FileReader();
    fr.onload = () => {
      try {
        const p = JSON.parse(fr.result);
        if (!p.items) throw new Error();
        setData({ prices: p.prices || {}, items: p.items });
        setStatus('불러왔습니다');
      } catch {
        setStatus('읽을 수 없는 파일입니다');
      }
      setTimeout(() => setStatus(''), 2000);
    };
    fr.readAsText(f);
    e.target.value = '';
  };

  const matNames = Object.keys(prices).sort((a, b) => a.localeCompare(b, 'ko'));

  if (!data) {
    return (
      <>
        <Style />
        <div className="page">
          <div className="loading">장부를 펴는 중…</div>
        </div>
      </>
    );
  }

  return (
    <>
      <Style />
      <div className="page">
        <header className="top">
          <div className="brand">
            <span className="seal" aria-hidden="true">
              商
            </span>
            <div>
              <h1>제작 장부</h1>
              <p className="sub">거상 아이템 원가 계산</p>
            </div>
          </div>

          <div className="ledger-sum">
            <div className="sum-cell">
              <span className="sum-key">등록</span>
              <span className="sum-val">{items.length}건</span>
            </div>
            <div className="sum-cell">
              <span className="sum-key">제작비 합</span>
              <span className="sum-val">{comma(grand.cost)}</span>
              <span className="sum-kr">{korean(grand.cost)}</span>
            </div>
            {grand.sell > 0 && (
              <div className="sum-cell">
                <span className="sum-key">예상 순이익</span>
                <span className={`sum-val ${grand.profit >= 0 ? 'up' : 'down'}`}>
                  {grand.profit >= 0 ? '+' : ''}
                  {comma(grand.profit)}
                </span>
                <span className="sum-kr">{korean(grand.profit)}</span>
              </div>
            )}
          </div>
        </header>

        <nav className="tabs" role="tablist">
          <button role="tab" aria-selected={tab === 'items'} className={tab === 'items' ? 'on' : ''} onClick={() => setTab('items')}>
            아이템
          </button>
          <button role="tab" aria-selected={tab === 'prices'} className={tab === 'prices' ? 'on' : ''} onClick={() => setTab('prices')}>
            시세표 <span className="count">{matNames.length}</span>
          </button>
          <button role="tab" aria-selected={tab === 'sync'} className={tab === 'sync' ? 'on' : ''} onClick={() => setTab('sync')}>
            동기화 {linked && <span className="dot" aria-label="연결됨" />}
          </button>
          <div className="spacer" />
          <span className="status" aria-live="polite">
            {status}
          </span>
          <button className="ghost" onClick={toExcel}>
            엑셀로 받기
          </button>
          <button className="ghost" onClick={backup}>
            백업
          </button>
          <button className="ghost" onClick={() => fileRef.current?.click()}>
            복원
          </button>
          <input ref={fileRef} type="file" accept="application/json" onChange={restore} hidden />
        </nav>

        {tab !== 'sync' && (
        <div className="search">
          <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">
            <circle cx="7" cy="7" r="4.6" fill="none" stroke="currentColor" strokeWidth="1.6" />
            <path d="M10.5 10.5 L14 14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
          </svg>
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder={tab === 'items' ? '아이템 이름, 제작 경로, 재료 이름으로 찾기' : '재료 이름으로 찾기'}
            aria-label="검색"
            onKeyDown={(e) => {
              if (e.key === 'Escape') setQ('');
            }}
          />
          {q && (
            <>
              <span className="found">
                {tab === 'items' ? `${shownItems.length}건` : `${shownMats.length}개`}
              </span>
              <button className="x" aria-label="검색어 지우기" onClick={() => setQ('')}>
                ×
              </button>
            </>
          )}
        </div>
        )}

        {tab === 'items' ? (
          <main>
            {items.length === 0 && (
              <div className="empty">
                <p>장부가 비어 있습니다.</p>
                <button className="solid" onClick={addItem}>
                  첫 아이템 적기
                </button>
              </div>
            )}

            {items.length > 0 && shownItems.length === 0 && (
              <div className="empty">
                <p>‘{q}’와 맞는 아이템이 없습니다.</p>
                <button className="ghost" onClick={() => setQ('')}>
                  검색어 지우기
                </button>
              </div>
            )}

            {needle && shownItems.length > 0 && (
              <p className="reorder-note">검색 중에는 순서를 바꿀 수 없습니다. 검색어를 지우면 다시 됩니다.</p>
            )}

            {shownItems.map(({ it, hitMats }, idx) => {
              const c = calc(it);
              const open = needle ? true : it.open;
              const movable = !needle;
              return (
                <section
                  className={`item${hitMats.length ? ' matched' : ''}${dragId === it.id ? ' dragging' : ''}${
                    overId === it.id && dragId !== it.id ? ' over' : ''
                  }`}
                  key={it.id}
                  onDragStart={() => setDragId(it.id)}
                  onDragEnd={() => {
                    setDragId(null);
                    setOverId(null);
                  }}
                  onDragOver={(e) => {
                    if (dragId && dragId !== it.id) {
                      e.preventDefault();
                      setOverId(it.id);
                    }
                  }}
                  onDragLeave={() => setOverId((v) => (v === it.id ? null : v))}
                  onDrop={(e) => {
                    e.preventDefault();
                    dropOn(dragId, it.id);
                    setDragId(null);
                    setOverId(null);
                  }}
                >
                  <div className="item-head">
                    {movable && (
                      <span
                        className="handle"
                        draggable
                        role="button"
                        tabIndex={-1}
                        aria-label="끌어서 순서 바꾸기"
                        title="끌어서 순서를 바꿀 수 있습니다"
                      >
                        <svg viewBox="0 0 10 16" width="10" height="16" aria-hidden="true">
                          <g fill="currentColor">
                            <circle cx="3" cy="3" r="1.3" />
                            <circle cx="7" cy="3" r="1.3" />
                            <circle cx="3" cy="8" r="1.3" />
                            <circle cx="7" cy="8" r="1.3" />
                            <circle cx="3" cy="13" r="1.3" />
                            <circle cx="7" cy="13" r="1.3" />
                          </g>
                        </svg>
                      </span>
                    )}
                    <button
                      className="disc"
                      aria-expanded={!!open}
                      aria-label={open ? '접기' : '펼치기'}
                      onClick={() => patchItem(it.id, { open: !open })}
                    >
                      <svg viewBox="0 0 10 10" width="10" height="10" className={open ? 'rot' : ''}>
                        <path d="M2 1 L8 5 L2 9 Z" fill="currentColor" />
                      </svg>
                    </button>

                    <input
                      className="name"
                      value={it.name}
                      placeholder="아이템 이름"
                      aria-label="아이템 이름"
                      onChange={(e) => patchItem(it.id, { name: e.target.value })}
                    />
                    <input
                      className="path"
                      value={it.path}
                      placeholder="제작 경로"
                      aria-label="제작 경로"
                      onChange={(e) => patchItem(it.id, { path: e.target.value })}
                    />

                    <label className="count-box" title="한 번에 몇 개를 만들지">
                      <span className="times">×</span>
                      <input
                        className="cnt"
                        value={it.count === 0 ? '' : it.count ?? 1}
                        aria-label="제작 수량"
                        inputMode="numeric"
                        placeholder="1"
                        onChange={(e) => patchItem(it.id, { count: parseAmount(e.target.value) })}
                        onBlur={(e) => {
                          if (!parseAmount(e.target.value)) patchItem(it.id, { count: 1 });
                        }}
                      />
                      <span className="unit">개</span>
                    </label>

                    <div className="head-total">
                      <span className="ht-num">{comma(c.total)}</span>
                      <span className="ht-kr">{korean(c.total)}</span>
                    </div>

                    <div className="head-act">
                      {movable && (
                        <span className="nudge">
                          <button
                            className="ghost sm"
                            aria-label="위로 한 칸"
                            disabled={idx === 0}
                            onClick={() => moveItem(it.id, -1)}
                          >
                            ↑
                          </button>
                          <button
                            className="ghost sm"
                            aria-label="아래로 한 칸"
                            disabled={idx === shownItems.length - 1}
                            onClick={() => moveItem(it.id, 1)}
                          >
                            ↓
                          </button>
                        </span>
                      )}
                      <button className="ghost sm" onClick={() => dupItem(it)}>
                        복제
                      </button>
                      <button className="ghost sm danger" onClick={() => delItem(it.id)}>
                        삭제
                      </button>
                    </div>

                    {hitMats.length > 0 && (
                      <p className="hit-line">
                        걸린 재료 {hitMats.map((n) => (
                          <span className="hit-chip" key={n}>
                            {n}
                          </span>
                        ))}
                      </p>
                    )}
                  </div>

                  {open && (
                    <div className="body">
                      <div className="grid head-row">
                        <span>재료</span>
                        <span className="r">개당 가격</span>
                        <span className="r">갯수</span>
                        <span className="r">합계</span>
                        <span />
                      </div>

                      {it.mats.map((m) => {
                        const unit = prices[m.name] || 0;
                        const need = (m.qty || 0) * c.n;
                        const line = unit * need;
                        const shared = m.name && (usage[m.name] || 0) > 1;
                        const hit = needle && m.name && m.name.toLowerCase().includes(needle);
                        return (
                          <div className={`grid mat-row${hit ? ' hit' : ''}`} key={m.id}>
                            <div className="mat-name">
                              <input
                                list="matlist"
                                value={m.name}
                                placeholder="재료 이름"
                                aria-label="재료 이름"
                                onChange={(e) => patchMat(it.id, m.id, { name: e.target.value })}
                              />
                              {shared && <span className="shared" title={`${usage[m.name]}개 아이템이 이 재료를 씁니다`}>공용 {usage[m.name]}</span>}
                            </div>

                            <AmountInput label="개당 가격" value={unit} placeholder="0" onCommit={(v) => setPrice(m.name, v)} />

                            <div className="qty-cell">
                              <input
                                className="amt qty"
                                value={m.qty ?? ''}
                                aria-label="1개당 갯수"
                                inputMode="numeric"
                                onChange={(e) => patchMat(it.id, m.id, { qty: parseAmount(e.target.value) })}
                              />
                              {c.n > 1 && <span className="need">{comma(need)}개 필요</span>}
                            </div>

                            <div className="line">
                              <span className="line-num">{comma(line)}</span>
                              <span className="line-kr">{korean(line)}</span>
                            </div>

                            <button className="x" aria-label="재료 삭제" onClick={() => delMat(it.id, m.id)}>
                              ×
                            </button>
                          </div>
                        );
                      })}

                      <button className="add" onClick={() => addMat(it.id)}>
                        재료 한 줄 더
                      </button>

                      <div className="grid fee-row">
                        <span className="fee-label">수수료 {c.n > 1 && <em>1회당</em>}</span>
                        <AmountInput label="수수료" value={it.fee} placeholder="0" onCommit={(v) => patchItem(it.id, { fee: v })} />
                        <span />
                        <div className="line">
                          <span className="line-num">{comma(c.fee)}</span>
                          <span className="line-kr">{korean(c.fee)}</span>
                        </div>
                        <span />
                      </div>

                      <div className="total-row">
                        <span className="total-label">
                          총 제작비 {c.n > 1 && <em>{comma(c.n)}개분</em>}
                        </span>
                        <div className="total-num">
                          <span className="big">{comma(c.total)}</span>
                          <span className="big-kr">
                            {korean(c.total)}
                            {c.n > 1 && ` · 1개당 ${comma(c.unitCost)}`}
                          </span>
                        </div>
                      </div>

                      <div className="sell-row">
                        <label>
                          <span>판매가 {c.n > 1 && <em>1개당</em>}</span>
                          <AmountInput
                            label="판매가"
                            value={it.sellPrice}
                            placeholder="팔 가격"
                            onCommit={(v) => patchItem(it.id, { sellPrice: v })}
                          />
                        </label>

                        {it.sellPrice ? (
                          <div className={`profit ${c.profit >= 0 ? 'up' : 'down'}`}>
                            <span className="p-num">
                              {c.profit >= 0 ? '+' : ''}
                              {comma(c.profit)}
                            </span>
                            <span className="p-meta">
                              {korean(c.profit)} · 마진 {(c.margin * 100).toFixed(1)}%
                              {c.n > 1 && ` · 1개당 ${c.profitUnit >= 0 ? '+' : ''}${comma(c.profitUnit)}`}
                            </span>
                          </div>
                        ) : (
                          <p className="hint">판매가를 적으면 순이익과 마진이 나옵니다.</p>
                        )}

                        <input
                          className="memo"
                          value={it.memo}
                          placeholder="메모"
                          aria-label="메모"
                          onChange={(e) => patchItem(it.id, { memo: e.target.value })}
                        />
                      </div>
                    </div>
                  )}
                </section>
              );
            })}

            {items.length > 0 && !needle && (
              <button className="solid wide" onClick={addItem}>
                아이템 추가
              </button>
            )}
          </main>
        ) : tab === 'prices' ? (
          <main>
            <p className="lead">
              재료 가격은 여기 한 곳에만 저장됩니다. 값을 고치면 그 재료를 쓰는 아이템의 원가가 전부 다시 계산됩니다.
            </p>

            {matNames.length === 0 && <div className="empty">아직 등록된 재료가 없습니다.</div>}

            {matNames.length > 0 && shownMats.length === 0 && (
              <div className="empty">
                <p>‘{q}’와 맞는 재료가 없습니다.</p>
                <button className="ghost" onClick={() => setQ('')}>
                  검색어 지우기
                </button>
              </div>
            )}

            {shownMats.map((n) => (
              <div className="price-row" key={n}>
                <span className="p-name">
                  <Mark text={n} q={needle} />
                </span>
                <span className="p-use">{usage[n] ? `아이템 ${usage[n]}건` : '쓰는 곳 없음'}</span>
                <AmountInput label={`${n} 가격`} value={prices[n]} onCommit={(v) => setPrice(n, v)} />
                <span className="p-kr">{korean(prices[n])}</span>
                {!usage[n] && (
                  <button
                    className="x"
                    aria-label={`${n} 삭제`}
                    onClick={() =>
                      setData((d) => {
                        const p = { ...d.prices };
                        delete p[n];
                        return { ...d, prices: p };
                      })
                    }
                  >
                    ×
                  </button>
                )}
              </div>
            ))}

            <div className="reset">
              {confirmReset ? (
                <>
                  <span>장부를 전부 지웁니다. 되돌릴 수 없습니다.</span>
                  <button
                    className="ghost sm danger"
                    onClick={() => {
                      setData({ prices: {}, items: [] });
                      setConfirmReset(false);
                    }}
                  >
                    지웁니다
                  </button>
                  <button className="ghost sm" onClick={() => setConfirmReset(false)}>
                    그만두기
                  </button>
                </>
              ) : (
                <button className="ghost sm danger" onClick={() => setConfirmReset(true)}>
                  장부 전체 지우기
                </button>
              )}
            </div>
          </main>
        ) : (
          <main>
            {!linked ? (
              <div className="sync-box">
                <h2>기기끼리 장부 맞추기</h2>
                <p className="lead">
                  GitHub의 비공개 Gist 한 칸을 보관함으로 씁니다. 토큰을 한 번 넣어두면 이 기기에서 고친 내용이 자동으로
                  올라가고, 다른 기기에서 열면 최신 장부를 가져옵니다.
                </p>

                <ol className="howto">
                  <li>
                    <a href="https://github.com/settings/tokens">github.com/settings/tokens</a> 로 갑니다.
                  </li>
                  <li>Tokens (classic) 을 고르고 Generate new token (classic) 을 누릅니다.</li>
                  <li>
                    Note 에 아무 이름이나 적고, Expiration 에서 기간을 정합니다. 아래 권한 목록에서 <b>gist 하나만</b>{' '}
                    체크하세요.
                  </li>
                  <li>Generate token 을 누르면 한 번만 보이는 문자열이 나옵니다. 복사해서 아래에 붙여넣으세요.</li>
                </ol>

                <label className="field">
                  <span>토큰</span>
                  <input
                    type="password"
                    value={tokenDraft}
                    placeholder="ghp_ 로 시작하는 문자열"
                    onChange={(e) => setTokenDraft(e.target.value)}
                  />
                </label>

                <div className="row">
                  <button className="solid" disabled={busy} onClick={() => connect('new')}>
                    {busy ? '연결하는 중…' : '보관함 새로 만들기'}
                  </button>
                  <span className="hint">이 기기가 처음이라면 이쪽입니다.</span>
                </div>

                <div className="or">이미 다른 기기에서 만들어 두셨다면</div>

                <label className="field">
                  <span>Gist 주소</span>
                  <input
                    value={gistDraft}
                    placeholder="https://gist.github.com/... 또는 아이디"
                    onChange={(e) => setGistDraft(e.target.value)}
                  />
                </label>
                <div className="row">
                  <button className="ghost" disabled={busy} onClick={() => connect('existing')}>
                    기존 보관함에 연결
                  </button>
                  <span className="hint">그 기기의 장부를 그대로 가져옵니다.</span>
                </div>

                {syncMsg && <p className="sync-msg">{syncMsg}</p>}

                <p className="warn">
                  토큰은 이 브라우저에만 저장됩니다. 권한을 gist 하나로 좁혀두면, 새어 나가더라도 남이 건드릴 수 있는 건
                  이 장부뿐입니다. 공용 PC에서는 다 쓰신 뒤 연결을 끊어주세요.
                </p>
              </div>
            ) : (
              <div className="sync-box">
                <h2>연결됨</h2>
                <p className="lead">고칠 때마다 자동으로 올라갑니다. 다른 기기에서 여시면 최신 장부를 가져옵니다.</p>

                <div className="field ro">
                  <span>보관함</span>
                  <a href={`https://gist.github.com/${sync.gistId}`}>{sync.gistId}</a>
                </div>

                <div className="row">
                  <button className="ghost" disabled={busy} onClick={pullNow}>
                    지금 가져오기
                  </button>
                  <button className="ghost" disabled={busy} onClick={pushNow}>
                    지금 올리기
                  </button>
                  <button className="ghost danger" disabled={busy} onClick={unlink}>
                    연결 끊기
                  </button>
                </div>

                {syncMsg && <p className="sync-msg">{syncMsg}</p>}

                <p className="warn">
                  두 기기에서 동시에 고치면 나중에 올라간 쪽이 남습니다. 다른 기기에서 작업하다 왔다면 먼저 가져오기를
                  누르세요.
                </p>

                <div className="row" style={{ marginTop: '18px' }}>
                  <button className="ghost sm" onClick={() => setTab('items')}>
                    장부로 돌아가기
                  </button>
                </div>
              </div>
            )}
          </main>
        )}

        <datalist id="matlist">
          {matNames.map((n) => (
            <option key={n} value={n} />
          ))}
        </datalist>

        <footer className="foot">
          입력한 내용은 이 브라우저에 자동으로 남습니다. 다른 기기로 옮기려면 백업 파일을 쓰세요.
        </footer>
      </div>
    </>
  );
}

/* 검색어와 겹치는 부분만 색을 준다 */
function Mark({ text, q }) {
  if (!q) return <>{text}</>;
  const i = (text || '').toLowerCase().indexOf(q);
  if (i < 0) return <>{text}</>;
  return (
    <>
      {text.slice(0, i)}
      <mark>{text.slice(i, i + q.length)}</mark>
      {text.slice(i + q.length)}
    </>
  );
}

/* ────────────────────────────────────────────────────────────
   백자 바탕 · 청화 코발트 · 인주 붉은색
   ──────────────────────────────────────────────────────────── */

function Style() {
  return (
    <style>{`
@import url('https://fonts.googleapis.com/css2?family=Gowun+Batang:wght@400;700&family=IBM+Plex+Sans+KR:wght@400;500;600&display=swap');

.page{
  --paper:#E3E7DF;
  --card:#F8F9F5;
  --rule:#C6CCC0;
  --rule-soft:#DDE2D8;
  --ink:#23231F;
  --ink-soft:#6C7067;
  --cobalt:#1E3F72;
  --cobalt-mid:#40639A;
  --cinnabar:#A6392C;
  --celadon:#3F7255;
  --brass:#8A6E3B;

  background:var(--paper);
  color:var(--ink);
  font-family:'IBM Plex Sans KR',system-ui,sans-serif;
  font-size:14px;
  line-height:1.5;
  min-height:100vh;
  padding:22px 18px 40px;
  box-sizing:border-box;
}
.page *{box-sizing:border-box;}
.page :focus-visible{outline:2px solid var(--cobalt);outline-offset:1px;}

.loading{padding:60px;text-align:center;color:var(--ink-soft);font-family:'Gowun Batang',serif;}

/* 머리 */
.top{
  max-width:1000px;margin:0 auto 18px;
  display:flex;flex-wrap:wrap;gap:20px;align-items:flex-end;justify-content:space-between;
  border-bottom:2px solid var(--cobalt);padding-bottom:14px;
}
.brand{display:flex;gap:12px;align-items:center;}
.seal{
  width:38px;height:38px;flex:none;border-radius:3px;
  background:var(--cinnabar);color:#F8F4EC;
  display:grid;place-items:center;
  font-family:'Gowun Batang',serif;font-size:21px;font-weight:700;
}
.top h1{
  font-family:'Gowun Batang',serif;font-size:26px;font-weight:700;
  margin:0;letter-spacing:.06em;color:var(--cobalt);
}
.sub{margin:1px 0 0;font-size:12px;color:var(--ink-soft);}

.ledger-sum{display:flex;gap:26px;flex-wrap:wrap;}
.sum-cell{display:flex;flex-direction:column;align-items:flex-end;min-width:96px;}
.sum-key{font-size:11px;color:var(--ink-soft);}
.sum-val{
  font-size:19px;font-weight:600;font-variant-numeric:tabular-nums;
  letter-spacing:-.01em;color:var(--cobalt);
}
.sum-kr{font-size:11px;color:var(--ink-soft);font-family:'Gowun Batang',serif;}
.up{color:var(--celadon);} .down{color:var(--cinnabar);}

/* 탭 */
.tabs{
  max-width:1000px;margin:0 auto 14px;
  display:flex;gap:6px;align-items:center;flex-wrap:wrap;
}
.tabs [role=tab]{
  font-family:inherit;font-size:14px;padding:7px 15px;
  border:1px solid var(--rule);border-bottom:none;
  border-radius:3px 3px 0 0;background:transparent;color:var(--ink-soft);
  cursor:pointer;
}
.tabs [role=tab].on{background:var(--card);color:var(--cobalt);font-weight:600;border-color:var(--cobalt);}
.count{font-size:11px;color:var(--ink-soft);margin-left:3px;}
.spacer{flex:1;}
.status{font-size:12px;color:var(--celadon);margin-right:6px;}

/* 검색 */
.search{
  max-width:1000px;margin:0 auto 14px;
  display:flex;align-items:center;gap:8px;
  background:var(--card);border:1px solid var(--rule);border-radius:3px;
  padding:7px 11px;color:var(--ink-soft);
}
.search:focus-within{border-color:var(--cobalt);color:var(--cobalt);}
.search input{
  flex:1;min-width:0;background:transparent;border:none;outline:none;
  font-family:inherit;font-size:14px;color:var(--ink);padding:1px 0;
}
.search input::placeholder{color:#A9AEA2;}
.found{font-size:12px;color:var(--cobalt);font-variant-numeric:tabular-nums;flex:none;}

mark{background:#F2E3A8;color:var(--ink);border-radius:2px;padding:0 1px;}

.item.matched{border-color:var(--cobalt-mid);}

/* 순서 바꾸기 */
.handle{
  color:#B4B9AD;cursor:grab;padding:2px 1px;line-height:0;flex:none;touch-action:none;
}
.handle:hover{color:var(--cobalt);}
.handle:active{cursor:grabbing;}
.item.dragging{opacity:.45;}
.item.over{border-color:var(--cobalt);box-shadow:inset 0 3px 0 -1px var(--cobalt);}
.nudge{display:inline-flex;gap:2px;margin-right:3px;}
.nudge button{padding:3px 7px;line-height:1.1;font-size:13px;}
.reorder-note{
  font-size:12px;color:var(--ink-soft);margin:0 0 10px;
}
.hit-line{
  flex-basis:100%;margin:2px 0 0 30px;font-size:11px;color:var(--ink-soft);
  display:flex;gap:5px;align-items:center;flex-wrap:wrap;
}
.hit-chip{
  background:#F2E3A8;border-radius:2px;padding:1px 6px;color:var(--ink);font-size:11px;
}
.mat-row.hit{background:#FBF6E4;margin:0 -7px;padding-left:7px;padding-right:7px;}

button{font-family:inherit;cursor:pointer;}
.ghost{
  background:transparent;border:1px solid var(--rule);border-radius:3px;
  padding:6px 12px;font-size:13px;color:var(--ink-soft);
}
.ghost:hover{border-color:var(--cobalt);color:var(--cobalt);}
.ghost.sm{padding:4px 9px;font-size:12px;}
.ghost.danger:hover{border-color:var(--cinnabar);color:var(--cinnabar);}
.solid{
  background:var(--cobalt);color:#F4F6F1;border:none;border-radius:3px;
  padding:9px 20px;font-size:14px;
}
.solid:hover{background:var(--cobalt-mid);}
.wide{display:block;max-width:1000px;margin:6px auto 0;width:100%;}

main{max-width:1000px;margin:0 auto;}
.lead{font-size:13px;color:var(--ink-soft);margin:0 0 14px;max-width:62ch;}
.empty{
  background:var(--card);border:1px solid var(--rule);border-radius:3px;
  padding:34px;text-align:center;color:var(--ink-soft);
}
.empty p{margin:0 0 14px;font-family:'Gowun Batang',serif;font-size:15px;}

/* 아이템 */
.item{
  background:var(--card);border:1px solid var(--rule);border-radius:3px;
  margin-bottom:12px;
}
.item-head{
  display:flex;gap:10px;align-items:center;padding:11px 13px;flex-wrap:wrap;
}
.disc{
  background:none;border:none;color:var(--cobalt);padding:4px;line-height:0;flex:none;
}
.disc svg{transition:transform .16s ease;}
.disc .rot{transform:rotate(90deg);}
@media (prefers-reduced-motion:reduce){.disc svg{transition:none;}}

.item input, .price-row input{
  font-family:inherit;font-size:14px;color:var(--ink);
  background:transparent;border:none;border-bottom:1px solid var(--rule-soft);
  padding:4px 2px;
}
.item input:focus, .price-row input:focus{border-bottom-color:var(--cobalt);outline:none;}
.item input::placeholder{color:#A9AEA2;}

.name{
  font-family:'Gowun Batang',serif;font-size:18px !important;font-weight:700;
  color:var(--cobalt) !important;flex:1;min-width:150px;
}
.path{width:110px;font-size:13px !important;color:var(--ink-soft) !important;}

/* 제작 수량 */
.count-box{
  display:inline-flex;align-items:center;gap:3px;flex:none;
  border:1px solid var(--rule);border-radius:3px;padding:2px 7px;
}
.count-box:focus-within{border-color:var(--cobalt);}
.count-box .times{color:var(--ink-soft);font-size:13px;}
.count-box .unit{color:var(--ink-soft);font-size:12px;}
.item .cnt{
  width:42px;text-align:center;border-bottom:none !important;
  font-variant-numeric:tabular-nums;font-weight:600;color:var(--cobalt) !important;
  padding:2px 0 !important;
}
.qty-cell{display:flex;flex-direction:column;align-items:stretch;}
.need{
  font-size:10px;color:var(--brass);text-align:right;margin-top:1px;
  font-variant-numeric:tabular-nums;white-space:nowrap;
}
.fee-label em, .total-label em, .sell-row label em{
  font-style:normal;font-size:11px;color:var(--ink-soft);
  border:1px solid var(--rule);border-radius:2px;padding:0 4px;margin-left:4px;
}
.total-label em{border-color:var(--cobalt-mid);color:var(--cobalt-mid);}

.head-total{display:flex;flex-direction:column;align-items:flex-end;margin-left:auto;}
.ht-num{font-size:17px;font-weight:600;font-variant-numeric:tabular-nums;}
.ht-kr{font-size:11px;color:var(--ink-soft);font-family:'Gowun Batang',serif;}
.head-act{display:flex;gap:5px;}

.body{border-top:1px solid var(--rule);padding:12px 13px 14px;}

.grid{
  display:grid;
  grid-template-columns:minmax(140px,1.5fr) 130px 66px minmax(120px,1fr) 26px;
  gap:10px;align-items:center;
}
.head-row{font-size:11px;color:var(--ink-soft);padding-bottom:5px;border-bottom:1px solid var(--rule);}
.r{text-align:right;}

.mat-row{padding:6px 0;border-bottom:1px solid var(--rule-soft);}
.mat-name{display:flex;align-items:center;gap:6px;}
.mat-name input{flex:1;min-width:0;}
.shared{
  font-size:10px;color:var(--brass);border:1px solid var(--brass);
  border-radius:2px;padding:0 4px;flex:none;
}
.amt{
  width:100%;text-align:right;font-variant-numeric:tabular-nums;
}
.qty{width:100%;}
.line{display:flex;flex-direction:column;align-items:flex-end;}
.line-num{font-variant-numeric:tabular-nums;}
.line-kr{font-size:10px;color:var(--ink-soft);font-family:'Gowun Batang',serif;}
.x{
  background:none;border:none;color:#B4B9AD;font-size:17px;line-height:1;padding:2px 5px;
}
.x:hover{color:var(--cinnabar);}

.add{
  background:none;border:1px dashed var(--rule);border-radius:3px;
  color:var(--ink-soft);font-size:12px;padding:5px 12px;margin:9px 0 4px;
}
.add:hover{border-color:var(--cobalt);color:var(--cobalt);}

.fee-row{padding:6px 0;border-top:1px solid var(--rule);}
.fee-label{font-size:13px;color:var(--ink-soft);}

.total-row{
  display:flex;justify-content:space-between;align-items:baseline;
  border-top:2px solid var(--cobalt);margin-top:8px;padding-top:9px;
}
.total-label{font-family:'Gowun Batang',serif;font-size:15px;color:var(--cobalt);}
.total-num{display:flex;flex-direction:column;align-items:flex-end;}
.big{
  font-size:27px;font-weight:600;font-variant-numeric:tabular-nums;
  letter-spacing:-.015em;color:var(--cobalt);line-height:1.15;
}
.big-kr{font-family:'Gowun Batang',serif;font-size:13px;color:var(--ink-soft);}

.sell-row{
  display:flex;gap:20px;align-items:center;flex-wrap:wrap;
  margin-top:12px;padding-top:11px;border-top:1px solid var(--rule-soft);
}
.sell-row label{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink-soft);}
.sell-row label .amt{width:130px;}
.profit{display:flex;flex-direction:column;}
.p-num{font-size:17px;font-weight:600;font-variant-numeric:tabular-nums;}
.p-meta{font-size:11px;color:var(--ink-soft);}
.hint{font-size:12px;color:var(--ink-soft);margin:0;}
.memo{flex:1;min-width:130px;font-size:13px !important;}

/* 시세표 */
.price-row{
  display:grid;grid-template-columns:minmax(120px,1fr) 100px 140px 120px 26px;
  gap:12px;align-items:center;
  background:var(--card);border:1px solid var(--rule);border-top:none;padding:9px 13px;
}
.price-row:first-of-type{border-top:1px solid var(--rule);border-radius:3px 3px 0 0;}
.p-name{font-family:'Gowun Batang',serif;font-size:15px;}
.p-use{font-size:11px;color:var(--ink-soft);}
.p-kr{font-size:12px;color:var(--ink-soft);text-align:right;font-family:'Gowun Batang',serif;}
.reset{margin-top:20px;display:flex;gap:8px;align-items:center;font-size:12px;color:var(--ink-soft);}

/* 동기화 */
.dot{
  display:inline-block;width:6px;height:6px;border-radius:50%;
  background:var(--celadon);vertical-align:middle;margin-left:4px;
}
.sync-box{
  background:var(--card);border:1px solid var(--rule);border-radius:3px;
  padding:20px 22px;max-width:620px;
}
.sync-box h2{
  font-family:'Gowun Batang',serif;font-size:18px;font-weight:700;
  color:var(--cobalt);margin:0 0 8px;
}
.howto{margin:0 0 18px;padding-left:20px;font-size:13px;line-height:1.9;color:var(--ink);}
.howto li{padding-left:2px;}
.howto a{color:var(--cobalt);}
.howto b{font-weight:600;color:var(--cinnabar);}
.field{display:flex;align-items:center;gap:12px;margin-bottom:10px;}
.field>span{font-size:13px;color:var(--ink-soft);width:66px;flex:none;}
.field input{
  flex:1;min-width:0;font-family:inherit;font-size:13px;color:var(--ink);
  background:transparent;border:1px solid var(--rule);border-radius:3px;padding:7px 10px;
}
.field input:focus{border-color:var(--cobalt);outline:none;}
.field.ro a{font-size:13px;color:var(--cobalt);word-break:break-all;}
.row{display:flex;gap:9px;align-items:center;flex-wrap:wrap;margin-bottom:6px;}
.row .hint{font-size:12px;color:var(--ink-soft);}
.or{
  margin:22px 0 12px;font-size:12px;color:var(--ink-soft);
  border-top:1px solid var(--rule-soft);padding-top:14px;
}
.sync-msg{
  font-size:13px;color:var(--cobalt);background:var(--paper);
  border-radius:3px;padding:8px 11px;margin:12px 0 0;
}
.warn{
  font-size:12px;color:var(--ink-soft);line-height:1.7;
  border-left:2px solid var(--brass);padding-left:11px;margin:18px 0 0;border-radius:0;
}
button:disabled{opacity:.5;cursor:default;}

.foot{
  max-width:1000px;margin:26px auto 0;font-size:11px;color:var(--ink-soft);
  border-top:1px solid var(--rule);padding-top:10px;
}

@media (max-width:700px){
  .page{padding:16px 12px 34px;}
  .grid{grid-template-columns:1fr 84px 52px;grid-template-areas:'nm nm nm' 'pr qt ln';row-gap:4px;}
  .head-row{display:none;}
  .mat-row .mat-name{grid-area:nm;}
  .mat-row .amt:not(.qty){grid-area:pr;}
  .mat-row .qty{grid-area:qt;}
  .mat-row .line{grid-area:ln;}
  .mat-row .x{position:absolute;right:14px;margin-top:-26px;}
  .mat-row{position:relative;}
  .fee-row{grid-template-columns:1fr 100px;grid-template-areas:'fl pr' 'ln ln';}
  .fee-row .fee-label{grid-area:fl;}
  .fee-row .amt{grid-area:pr;}
  .fee-row .line{grid-area:ln;align-items:flex-start;}
  .price-row{grid-template-columns:1fr 120px;grid-template-areas:'nm pr' 'us kr';row-gap:3px;}
  .price-row .p-name{grid-area:nm;} .price-row .p-use{grid-area:us;}
  .price-row .amt{grid-area:pr;} .price-row .p-kr{grid-area:kr;text-align:left;}
  .price-row .x{display:none;}
  .mat-row .qty-cell{grid-area:qt;}
  .count-box{margin-left:auto;}
  .head-total{margin-left:0;}
  .big{font-size:23px;}
}
`}</style>
  );
}
