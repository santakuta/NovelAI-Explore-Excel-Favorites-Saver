// ==UserScript==
// @name         NovelAI Explore → Excel お気に入り保存
// @namespace    https://github.com/santakuta/NovelAI-Explore-Excel-Favorites-Saver
// @homepageURL  https://github.com/santakuta/NovelAI-Explore-Excel-Favorites-Saver
// @supportURL   https://github.com/santakuta/NovelAI-Explore-Excel-Favorites-Saver/issues
// @version      1.0.1
// @description  NovelAI Explore の画像ポップアップ（タイトル・プロンプト・設定値・サムネイル）をお気に入りとして蓄積し、1つのExcelファイルにまとめてエクスポートします。（非公式・NovelAIとは無関係）
// @license      MIT
// @match        https://novelai.net/explore/*
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_xmlhttpRequest
// @grant        GM_registerMenuCommand
// @connect      explore.novelai.net
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const STORE_KEY = 'nai_fav_items_v1';
  const API = 'https://explore.novelai.net';

  // ---------- storage ----------
  const load = () => { try { return JSON.parse(GM_getValue(STORE_KEY, '[]')); } catch { return []; } };
  const save = (items) => GM_setValue(STORE_KEY, JSON.stringify(items));

  // ---------- helpers ----------
  const currentPostId = () => {
    const m = location.pathname.match(/\/explore\/image\/([0-9a-f-]{36})/i);
    return m ? m[1] : null;
  };

  const gmGet = (url, responseType = 'json') => new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      method: 'GET', url, responseType,
      onload: (r) => (r.status >= 200 && r.status < 300) ? resolve(r.response) : reject(new Error(`HTTP ${r.status}`)),
      onerror: () => reject(new Error('network error')),
    });
  });

  const fmtDate = (iso) => {
    if (!iso) return '';
    const d = new Date(iso);
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  };

  function parsePost(j) {
    let meta = {}, c = {};
    try { meta = JSON.parse(j.image?.nai_metadata || '{}'); } catch {}
    try { c = typeof meta.Comment === 'string' ? JSON.parse(meta.Comment) : (meta.Comment || {}); } catch {}
    const v4 = c.v4_prompt?.caption || {};
    const v4n = c.v4_negative_prompt?.caption || {};
    const chars = (v4.char_captions || []).map((x) => x.char_caption).filter(Boolean);
    const charsNeg = (v4n.char_captions || []).map((x) => x.char_caption).filter(Boolean);
    return {
      id: j.id,
      title: j.title || '',
      creator: j.creator?.name || '',
      createdAt: j.created_at || '',
      description: j.description || '',
      likes: j.like_count ?? '',
      model: meta.Source || '',
      prompt: v4.base_caption || c.prompt || '',
      charPrompts: chars,
      negative: v4n.base_caption || c.uc || '',
      charNegatives: charsNeg,
      steps: c.steps ?? '',
      scale: c.scale ?? '',
      cfgRescale: c.cfg_rescale ?? '',
      seed: c.seed ?? '',
      sampler: c.sampler || '',
      noiseSchedule: c.noise_schedule || '',
      width: c.width || j.image?.width || '',
      height: c.height || j.image?.height || '',
      url: `https://novelai.net/explore/image/${j.id}`,
      savedAt: new Date().toISOString(),
      memo: '',
    };
  }

  // ---------- UI ----------
  const css = `
    #naifav-bar{position:fixed;right:18px;bottom:18px;z-index:2147483646;display:flex;gap:8px;font-family:sans-serif}
    #naifav-bar button{border:0;border-radius:8px;padding:9px 14px;font-size:13px;font-weight:bold;cursor:pointer;
      box-shadow:0 2px 8px rgba(0,0,0,.4);color:#1a1b2e;background:#f5f3c2}
    #naifav-bar button.sec{background:#2b2d4a;color:#f5f3c2}
    #naifav-bar button.saved{background:#7bd88f}
    #naifav-toast{position:fixed;right:18px;bottom:70px;z-index:2147483647;background:#2b2d4a;color:#fff;
      padding:10px 14px;border-radius:8px;font:13px sans-serif;opacity:0;transition:opacity .2s;pointer-events:none}
    #naifav-panel{position:fixed;inset:0;z-index:2147483647;background:rgba(0,0,0,.6);display:flex;align-items:center;justify-content:center}
    #naifav-panel .box{background:#1a1b2e;color:#eee;width:min(720px,92vw);max-height:80vh;border-radius:12px;padding:16px;
      display:flex;flex-direction:column;font:13px sans-serif}
    #naifav-panel .list{overflow:auto;flex:1;margin:10px 0}
    #naifav-panel .row{display:flex;gap:8px;align-items:center;padding:6px 0;border-bottom:1px solid #333}
    #naifav-panel .row a{color:#f5f3c2;flex:1;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    #naifav-panel button{border:0;border-radius:6px;padding:6px 10px;cursor:pointer;background:#2b2d4a;color:#f5f3c2}
    #naifav-panel .acts{display:flex;gap:8px;flex-wrap:wrap}
  `;
  const style = document.createElement('style');
  style.textContent = css;

  const bar = document.createElement('div');
  bar.id = 'naifav-bar';
  bar.innerHTML = `
    <button id="naifav-save" title="Alt+S">★ Excel用に保存</button>
    <button id="naifav-export" class="sec">⬇ Excel出力 (<span id="naifav-count">0</span>)</button>
    <button id="naifav-list" class="sec">一覧</button>`;
  const toastEl = document.createElement('div');
  toastEl.id = 'naifav-toast';

  // React(Next.js) が <body> を描き直すと追加した要素が消えるため、
  // <html> 直下に置き、消えていたら付け直す
  function ensureUI() {
    const root = document.documentElement;
    for (const el of [style, bar, toastEl]) {
      if (!el.isConnected) root.appendChild(el);
    }
  }
  ensureUI();

  const $ = (id) => bar.querySelector('#' + id);
  let toastTimer;
  const toast = (msg) => {
    toastEl.textContent = msg; toastEl.style.opacity = '1';
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (toastEl.style.opacity = '0'), 2200);
  };

  function refreshUI() {
    const id = currentPostId();
    const items = load();
    $('naifav-count').textContent = items.length;
    const btn = $('naifav-save');
    btn.style.display = id ? '' : 'none';
    const saved = id && items.some((x) => x.id === id);
    btn.classList.toggle('saved', !!saved);
    btn.textContent = saved ? '✓ 保存済み（再取得）' : '★ Excel用に保存';
  }

  // SPA のURL変化を監視
  let lastPath = '';
  setInterval(() => {
    ensureUI();
    if (location.pathname !== lastPath) { lastPath = location.pathname; refreshUI(); }
  }, 400);

  // ---------- actions ----------
  async function saveCurrent() {
    const id = currentPostId();
    if (!id) return toast('画像のポップアップを開いてから保存してください');
    try {
      const j = await gmGet(`${API}/post/${id}`);
      const rec = parsePost(j);
      const items = load();
      const i = items.findIndex((x) => x.id === id);
      if (i >= 0) { rec.savedAt = items[i].savedAt; rec.memo = items[i].memo; items[i] = rec; }
      else items.push(rec);
      save(items);
      toast(`保存しました：「${rec.title || '(無題)'}」（計 ${items.length} 件）`);
      refreshUI();
    } catch (e) {
      toast('保存に失敗しました: ' + e.message);
    }
  }

  // サムネイルを取得して JPEG(base64) に変換（Excel は webp 非対応のため）
  async function fetchThumbJpeg(id, maxH = 200) {
    try {
      const blob = await gmGet(`${API}/post/thumbnail/${id}`, 'blob');
      const bmp = await createImageBitmap(blob);
      const scale = Math.min(1, maxH / bmp.height);
      const w = Math.round(bmp.width * scale), h = Math.round(bmp.height * scale);
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      cv.getContext('2d').drawImage(bmp, 0, 0, w, h);
      return { base64: cv.toDataURL('image/jpeg', 0.85), w, h };
    } catch { return null; }
  }

  async function exportExcel() {
    const items = load();
    if (!items.length) return toast('まだ保存したものがありません');
    toast(`Excel を作成中…（${items.length} 件・サムネ取得中）`);

    const wb = new ExcelJS.Workbook();
    wb.creator = 'NovelAI Explore Favorites';
    const ws = wb.addWorksheet('お気に入り', { views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }] });
    ws.columns = [
      { header: '画像', key: 'img', width: 30 },
      { header: 'タイトル', key: 'title', width: 22 },
      { header: '投稿者', key: 'creator', width: 14 },
      { header: '投稿日', key: 'createdAt', width: 17 },
      { header: 'プロンプト（ベース）', key: 'prompt', width: 55 },
      { header: 'キャラプロンプト', key: 'chars', width: 55 },
      { header: 'ネガティブ', key: 'negative', width: 40 },
      { header: 'キャラネガティブ', key: 'charNeg', width: 30 },
      { header: 'ステップ', key: 'steps', width: 8 },
      { header: 'ガイダンス', key: 'scale', width: 9 },
      { header: 'CFG Rescale', key: 'cfgRescale', width: 10 },
      { header: 'シード', key: 'seed', width: 13 },
      { header: 'サンプラー', key: 'sampler', width: 18 },
      { header: 'ノイズスケジュール', key: 'noise', width: 12 },
      { header: 'サイズ', key: 'size', width: 11 },
      { header: 'モデル', key: 'model', width: 24 },
      { header: 'いいね', key: 'likes', width: 7 },
      { header: 'URL', key: 'url', width: 18 },
      { header: '保存日時', key: 'savedAt', width: 17 },
      { header: 'メモ', key: 'memo', width: 30 },
    ];
    const header = ws.getRow(1);
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2B2D4A' } };
    header.alignment = { vertical: 'middle' };
    header.height = 22;
    ws.autoFilter = { from: 'A1', to: 'T1' };

    // サムネを並列取得（最大4並列）
    const thumbs = new Array(items.length);
    let next = 0;
    await Promise.all(Array.from({ length: 4 }, async () => {
      while (next < items.length) { const k = next++; thumbs[k] = await fetchThumbJpeg(items[k].id); }
    }));

    items.forEach((it, idx) => {
      const r = ws.addRow({
        title: it.title,
        creator: it.creator,
        createdAt: fmtDate(it.createdAt),
        prompt: it.prompt,
        chars: it.charPrompts.map((p, i) => `[キャラ${i + 1}] ${p}`).join('\n\n'),
        negative: it.negative,
        charNeg: it.charNegatives.map((p, i) => `[キャラ${i + 1}] ${p}`).join('\n\n'),
        steps: it.steps, scale: it.scale, cfgRescale: it.cfgRescale,
        seed: it.seed === '' ? '' : String(it.seed),
        sampler: it.sampler, noise: it.noiseSchedule,
        size: it.width && it.height ? `${it.width}×${it.height}` : '',
        model: it.model, likes: it.likes,
        url: { text: '開く', hyperlink: it.url },
        savedAt: fmtDate(it.savedAt),
        memo: it.memo,
      });
      r.height = 155; // pt
      r.alignment = { vertical: 'top', wrapText: true };
      r.getCell('url').font = { color: { argb: 'FF1F6FEB' }, underline: true };
      const t = thumbs[idx];
      if (t) {
        const imgId = wb.addImage({ base64: t.base64, extension: 'jpeg' });
        const maxW = 205, maxH = 200; // px
        const s = Math.min(maxW / t.w, maxH / t.h, 1);
        ws.addImage(imgId, {
          tl: { col: 0.05, row: idx + 1 + 0.03 },
          ext: { width: Math.round(t.w * s), height: Math.round(t.h * s) },
          hyperlinks: { hyperlink: it.url },
        });
      }
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const d = new Date();
    const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `NovelAI_お気に入り_${stamp}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    toast(`Excel を出力しました（${items.length} 件）`);
  }

  function downloadText(name, text, type) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }

  function openPanel() {
    const items = load().slice().reverse();
    const wrap = document.createElement('div');
    wrap.id = 'naifav-panel';
    wrap.innerHTML = `<div class="box">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <b>保存済み：${items.length} 件</b><button data-act="close">閉じる</button></div>
      <div class="list">${items.map((it) => `
        <div class="row" data-id="${it.id}">
          <a href="${it.url}">${escapeHtml(it.title || '(無題)')} <span style="color:#999">— ${escapeHtml(it.creator)}</span></a>
          <input placeholder="メモ" value="${escapeHtml(it.memo || '')}" data-act="memo"
            style="width:180px;background:#2b2d4a;border:0;color:#eee;padding:5px;border-radius:5px">
          <button data-act="del">削除</button>
        </div>`).join('') || '<div style="color:#999">まだありません</div>'}</div>
      <div class="acts">
        <button data-act="export">⬇ Excel出力</button>
        <button data-act="backup">JSONバックアップ</button>
        <button data-act="restore">JSONから復元</button>
        <button data-act="clear" style="margin-left:auto;color:#ff8a8a">全件削除</button>
      </div></div>`;
    document.documentElement.appendChild(wrap);
    panelEl = wrap;

    wrap.__onClick = (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (e.target === wrap || act === 'close') { wrap.remove(); return; }
      const row = e.target.closest('.row');
      if (act === 'del' && row) {
        save(load().filter((x) => x.id !== row.dataset.id)); row.remove(); refreshUI();
      } else if (act === 'export') {
        exportExcel();
      } else if (act === 'backup') {
        downloadText('novelai_favorites_backup.json', JSON.stringify(load(), null, 2), 'application/json');
      } else if (act === 'restore') {
        const inp = document.createElement('input');
        inp.type = 'file'; inp.accept = '.json,application/json';
        inp.onchange = async () => {
          try {
            const add = JSON.parse(await inp.files[0].text());
            const map = new Map(load().map((x) => [x.id, x]));
            add.forEach((x) => x && x.id && map.set(x.id, x));
            save([...map.values()]); wrap.remove(); refreshUI(); openPanel();
            toast('復元しました');
          } catch (err) { toast('復元に失敗: ' + err.message); }
        };
        inp.click();
      } else if (act === 'clear') {
        if (confirm('保存済みのデータをすべて削除します。よろしいですか？')) { save([]); wrap.remove(); refreshUI(); }
      }
    };
    wrap.addEventListener('change', (e) => {
      if (e.target.dataset?.act !== 'memo') return;
      const id = e.target.closest('.row').dataset.id;
      const all = load(); const it = all.find((x) => x.id === id);
      if (it) { it.memo = e.target.value; save(all); toast('メモを保存しました'); }
    });
  }

  const escapeHtml = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ---------- wire up ----------
  // NovelAI のポップアップは「外側クリックで閉じる」処理がクリックを横取りするため、
  // window のキャプチャ段階（ページ側より先）で自前UIへのイベントを受け取り、ページ側には渡さない。
  let panelEl = null;
  const inUI = (t) => t instanceof Node && (bar.contains(t) || (panelEl?.isConnected && panelEl.contains(t)));
  ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'touchstart', 'touchend',
   'focusin', 'keydown', 'keyup', 'keypress'].forEach((type) => {
    window.addEventListener(type, (e) => { if (inUI(e.target)) e.stopImmediatePropagation(); }, true);
  });
  // ポップアップのフォーカス固定がメモ欄からフォーカスを奪わないようにする
  window.addEventListener('focusout', (e) => { if (inUI(e.relatedTarget)) e.stopImmediatePropagation(); }, true);
  window.addEventListener('click', (e) => {
    if (!inUI(e.target)) return;
    e.stopImmediatePropagation();
    const btn = e.target.closest('button');
    if (btn?.id === 'naifav-save') saveCurrent();
    else if (btn?.id === 'naifav-export') exportExcel();
    else if (btn?.id === 'naifav-list') openPanel();
    else if (panelEl?.contains(e.target)) panelEl.__onClick(e);
  }, true);

  document.addEventListener('keydown', (e) => {
    if (e.altKey && e.code === 'KeyS' && !/INPUT|TEXTAREA/.test(document.activeElement?.tagName)) {
      e.preventDefault(); saveCurrent();
    }
  });
  GM_registerMenuCommand('Excel出力', exportExcel);
  GM_registerMenuCommand('保存一覧を開く', openPanel);

  refreshUI();
})();
