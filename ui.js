/* 卒論 赤ペンチェック：画面・赤ペンPDF */
(function () {
  'use strict';
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  const CHECKS = [
    ['paper', '用紙・余白', 'A4縦、余白は上下左右15mm'],
    ['columns', '段組み', 'タイトル〜キーワードは1段、本文は2段'],
    ['pages', 'ページ数', '4ページ以上の偶数（多くは4〜6ページ）'],
    ['leftover', '見本の文字', '見本の説明の枠・例文・「明朝9ポイント」などの書き込みを消す'],
    ['title', 'タイトル', '明朝・太字・12ポイント、中央'],
    ['name', '氏名', '「3C(12)氏名」の形、明朝・9ポイント'],
    ['abstract', 'アブストラクト', '350字以上400字以内、明朝・9ポイント'],
    ['keyword', 'キーワード', '「キーワード：〇〇、〇〇」、明朝・太字・9ポイント'],
    ['structure', '章立て', '章の見出しがあり、番号が順にそろっている。最後に「文　献」（章の名前・順番は自由）'],
    ['headfont', '見出しの文字', '明朝・太字・10ポイント'],
    ['bodyfont', '本文の文字', '明朝・9ポイント（文献も）'],
    ['indent', '段落の書き出し', '1マスあける。文の途中でEnterを押さない'],
    ['style', '文体', '「です・ます」を使わず「である」調で'],
    ['quote', '引用の書き方', '「」の最後に「。」を入れない。「」の中の「」は『』'],
    ['refs', '文献の書き方', '(1)著者（年）「論文名」,『書名』pp., 出版社. の形。Wikipediaは使わない。サイトは見た日も'],
    ['cite', '本文中の出典', '引用・参考にした所に文献番号を付ける'],
    ['figure', '図表', '図は下、表は上に番号と題。出典を書く'],
    ['color', '文字の色', 'すべて黒（色・蛍光ペンなし）'],
    ['talk', '言葉づかい（参考）', '話し言葉・「〜と思う」の多用など。ルール外の、もう一歩の助言'],
  ];
  const LABEL = Object.fromEntries(CHECKS.map((c) => [c[0], c[1]]));
  const PRIORITY = ['leftover', 'pages', 'abstract', 'columns', 'style', 'structure', 'refs', 'bodyfont', 'headfont', 'title', 'name', 'keyword', 'paper', 'cite', 'indent', 'figure', 'quote', 'color', 'talk'];
  const MARK = { ok: '◯', warn: '△', ng: '✕', hint: '・', none: '―' };

  function circled(n) {
    if (n <= 20) return String.fromCharCode(0x245f + n);
    if (n <= 35) return String.fromCharCode(0x3250 + n - 20);
    if (n <= 50) return String.fromCharCode(0x32b1 + n - 36);
    return '(' + n + ')';
  }
  const lvCls = (l) => (l === 'ng' ? '' : l === 'warn' ? 'w' : 'h');

  // ---------- 講評のまとめ ----------
  function summary(res) {
    const counts = { ng: 0, warn: 0, ok: 0 };
    for (const [k] of CHECKS) {
      const c = res.checks[k];
      if (!c || k === 'talk') continue;
      if (c.status === 'ng') counts.ng++; else if (c.status === 'warn') counts.warn++; else if (c.status === 'ok') counts.ok++;
    }
    const top = [];
    for (const lv of ['ng', 'warn']) {
      for (const k of PRIORITY) {
        if (top.length >= 3) break;
        if (top.some((t) => t.key === k)) continue;
        const list = res.issues.filter((x) => x.key === k && x.level === lv);
        if (!list.length) continue;
        const spanN = list.reduce((s, x) => s + (x.spans ? x.spans.length : 1), 0);
        top.push({ key: k, level: lv, msg: (list[0].msg.length > 40 ? list[0].msg.split('。')[0] : list[0].msg) + (list.length > 1 ? `（ほか${list.length - 1}か所）` : ''), fix: list[0].msg.length > 40 ? (list[0].fix || '').split('／')[0] : list[0].fix, nos: list.map((x) => x.no), spanN });
      }
    }
    return { counts, top };
  }

  function senseiComment(res, sm) {
    const c = res.checks, s = res.stats;
    const goods = [];
    if (c.abstract && c.abstract.status === 'ok') goods.push(`アブストラクトは${s.abstractLen}字で、字数どおりです。`);
    if (c.style && c.style.status === 'ok') goods.push('「である」調で書けています。');
    if (c.refs && c.refs.status === 'ok') goods.push('文献の書き方がそろっています。');
    if (c.structure && c.structure.status === 'ok') goods.push('章立てがフォーマットどおりです。');
    if (c.indent && c.indent.status === 'ok') goods.push('段落の書き出しもそろっています。');
    if (c.pages && c.pages.status === 'ok') goods.push('ページ数も条件を満たしています。');
    let tail;
    if (!sm.counts.ng && !sm.counts.warn) tail = '体裁はほぼ完ぺきです。あとは中身――問題提起とまとめが対応しているか、主張に根拠があるか――をもう一度読み直そう。';
    else if (!sm.counts.ng) tail = '大きな直しはありません。△のところを見直せば仕上がりです。';
    else if (sm.counts.ng <= 3) tail = '直すところはあと少し。まずは「まず直すこと」から片付けよう。';
    else tail = '直すところが多く見えますが、同じ直しのくり返しがほとんどです。「まず直すこと」の3つから一つずつ片付けよう。';
    return goods.slice(0, 2).join('') + (goods.length ? '<br>' : '') + tail;
  }

  // ---------- 本文の1段落を赤ペンつきで ----------
  function paraHtml(b, issues) {
    if (b.type === 'table') {
      return '<table class="tbl">' + b.rows.slice(0, 12).map((r) => '<tr>' + r.map((c) => `<td>${esc(c)}</td>`).join('') + '</tr>').join('') + '</table>' + (b.rows.length > 12 ? '<div class="muted">（表の続きは省略）</div>' : '');
    }
    const text = b.text;
    const ev = []; // {at, kind:'open'|'close'|'no'|'img', ...}
    const spans = [];
    issues.forEach((x) => (x.spans || []).forEach((s, i) => spans.push({ s: s[0], e: Math.max(s[1], s[0] + 1), x, last: i === x.spans.length - 1 })));
    spans.sort((a, b) => a.s - b.s || b.e - a.e);
    let lastEnd = -1;
    const pos = [];
    for (const sp of spans) {
      if (sp.s < lastEnd) { if (sp.last) pos.push({ at: lastEnd, no: sp.x }); continue; }
      pos.push({ at: sp.s, open: sp.x });
      pos.push({ at: Math.min(sp.e, text.length), close: sp.x, no: sp.last ? sp.x : null });
      lastEnd = sp.e;
    }
    b.images.forEach((im) => pos.push({ at: im.at, img: true }));
    pos.sort((a, b) => a.at - b.at || (a.close ? -1 : 0) - (b.close ? -1 : 0));
    let out = '', i = 0;
    const heads = issues.filter((x) => !x.spans || !x.spans.length);
    if (heads.length) out += heads.map((x) => `<span class="no ${lvCls(x.level)}">${circled(x.no)}</span>`).join('');
    if (b.autoNo) out += `<span style="color:#888">${b.autoNo}．</span>`;
    for (const p of pos) {
      if (p.at > i) { out += esc(text.slice(i, p.at)); i = p.at; }
      if (p.open) out += `<span class="mark ${lvCls(p.open.level)}">`;
      if (p.close) out += '</span>';
      if (p.no) out += `<span class="no ${lvCls(p.no.level)}">${circled(p.no.no)}</span>`;
      if (p.img) out += '<span class="fig">［図・画像］</span>';
    }
    out += esc(text.slice(i));
    out = out.replace(/\t/g, '　');
    if (b.textboxes.length) out += b.textboxes.map((t) => `<div class="muted" style="border:1px dashed #bbb;padding:2px 8px;margin:2px 0;font-size:12px">［枠］${esc(t)}</div>`).join('');
    return out || '&nbsp;';
  }

  function noteHtml(x) {
    return `<div class="note ${x.level}"><span class="nn">${circled(x.no)}</span>${esc(x.msg)}${x.fix ? `<span class="fx">${esc(x.fix)}</span>` : ''}</div>`;
  }

  // ---------- 赤ペンシート全体 ----------
  function reportHtml(res) {
    const d = res.doc, s = res.stats, sm = summary(res);
    const who = s.name || d.fileName;
    const today = new Date();
    const date = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
    let h = '<div class="sheet">';
    // --- 1枚目：講評 ---
    h += '<div class="blk">';
    h += `<div class="s-head"><h2>卒論 赤ペンシート</h2><div class="who">${esc(who)}<br>${esc(date)}</div></div>`;
    h += `<div class="s-title">${esc(s.title || '（タイトルなし）')}</div>`;
    h += `<div class="nums"><span>ページ <b>${s.pagesWord ? s.pagesWord : '約' + s.pages}</b>${s.pagesWord ? '' : '（Windows版Wordでの目安）'}</span><span>アブストラクト <b>${s.abstractLen}</b>字</span><span>本文 <b>${s.bodyChars}</b>字</span><span>章 <b>${(s.headings || []).length}</b></span><span>文献 <b>${s.refs || 0}</b>件</span><span>図 <b>${s.figures || 0}</b>・表 <b>${s.tables || 0}</b></span></div>`;
    h += `<div class="score"><div class="ng"><b>${sm.counts.ng}</b>✕ 直す</div><div class="warn"><b>${sm.counts.warn}</b>△ 見直す</div><div class="ok"><b>${sm.counts.ok}</b>◯ できている</div></div>`;
    h += '</div>';
    if (sm.top.length) {
      h += '<div class="blk"><div class="sec-h">まず直すこと</div><ol class="top3">';
      sm.top.forEach((t, i) => { h += `<li><span class="n">${i + 1}</span><div><div class="m"><span class="tag">${esc(LABEL[t.key])}</span>${esc(t.msg)}</div><div class="f">→ ${esc(t.fix || '')}　<span style="font-size:13px">（赤ペン ${t.nos.slice(0, 6).map(circled).join('')}${t.nos.length > 6 ? '…' : ''}）</span></div></div></li>`; });
      h += '</ol></div>';
    }
    h += '<div><div class="sec-h blk">チェック表</div><table class="chk">';
    for (const [k, lb] of CHECKS) {
      const c = res.checks[k] || { status: 'none', detail: '' };
      h += `<tr class="blk"><td class="mk ${c.status}">${MARK[c.status] || ''}</td><td class="lb">${esc(lb)}</td><td class="dt">${esc(c.detail)}</td></tr>`;
    }
    h += '</table></div>';
    h += `<div class="blk"><div class="sensei">${senseiComment(res, sm)}<div class="sig">赤ペン先生より</div></div></div>`;
    h += '<div class="pagebreak pb"></div>';
    // --- 2枚目から：本文に赤ペン ---
    h += '<div class="blk"><div class="s-head"><h2>本文への赤ペン</h2><div class="who">' + esc(who) + '</div></div>';
    h += '<p class="pp-note">読みやすいように1段に並べ直しています（実際のレイアウトはWordで確かめよう）。</p>';
    h += '<div class="legend"><span><i style="background:var(--red)"></i>✕ 直す</span><span><i style="background:var(--amber)"></i>△ 見直す</span><span><i style="background:var(--blue)"></i>・ もう一歩（参考）</span></div></div>';
    const whole = res.issues.filter((x) => x.block === null || x.block === undefined);
    if (whole.length) h += `<div class="blk row"><div class="txt has" style="font-family:'Zen Kaku Gothic New';color:#777;font-size:13px">［文書全体の設定］</div><div class="notes">${whole.map(noteHtml).join('')}</div></div>`;
    let hn = 0;
    for (const b of d.blocks) {
      const iss = res.issues.filter((x) => x.block === b.idx);
      if (b.role === 'heading') { hn++; if (b.numbered && !/^[\s　]*[0-9０-９]/.test(b.text)) b.autoNo = hn; }
      if (b.role === 'empty' && !iss.length) continue;
      const cls = iss.some((x) => x.level === 'ng') ? 'has' : iss.some((x) => x.level === 'warn') ? 'hasw' : '';
      const ind = b.role === 'body' && ((b.ind.firstLineChars || 0) > 0 || (b.ind.firstLine || 0) > 60) ? ` style="text-indent:${Math.max(1, Math.round((b.ind.firstLineChars || b.ind.firstLine / 18) / 100))}em"` : '';
      const align = ['title', 'subtitle', 'name', 'refhead'].includes(b.role) ? ` style="text-align:${b.jc === 'center' ? 'center' : b.jc === 'right' || b.jc === 'end' ? 'right' : 'left'}"` : '';
      h += `<div class="blk row"><div class="txt r-${b.role || 'other'} ${cls}"${ind}${align}>${paraHtml(b, iss)}</div><div class="notes">${iss.map(noteHtml).join('')}</div></div>`;
    }
    h += `<div class="blk foot">この赤ペンは「卒業論文 レイアウトフォーマット（2026）」と「卒業論文及びアブストラクトを書くにあたって」をもとに、プログラムが体裁だけを確かめたものです。内容のよしあしは判定していません。ファイル：${esc(d.fileName)}</div>`;
    h += '</div>';
    return h;
  }

  // ---------- PDF ----------
  async function sheetToPdf(sheetHtml) {
    const stage = $('#stage');
    stage.innerHTML = `<div class="pdf-mode">${sheetHtml}</div>`;
    await document.fonts.ready;
    await new Promise((r) => setTimeout(r, 60));
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: 'mm', format: 'a4', compress: true });
    const M = 12, W = 210 - M * 2, H = 297 - M * 2;
    // シート全体を1回で画像にし、段落の切れ目で切ってページに並べる
    const sheet = stage.querySelector('.sheet');
    const SCALE = Math.min(1.8, 30000 / Math.max(1, sheet.offsetHeight)); // ブラウザの画像の高さの上限対策
    const cv = await html2canvas(sheet, { scale: SCALE, backgroundColor: '#ffffff', logging: false });
    const top0 = sheet.getBoundingClientRect().top;
    const units = Array.from(sheet.querySelectorAll('.blk, .pb')).map((el) => {
      const r = el.getBoundingClientRect();
      return { pb: el.classList.contains('pb'), y: (r.top - top0) * SCALE, h: r.height * SCALE };
    });
    const pxPerMm = cv.width / W;
    let y = M, pages = 1;
    const place = (sy, sh) => {
      const c2 = document.createElement('canvas');
      c2.width = cv.width; c2.height = Math.max(1, Math.round(sh));
      c2.getContext('2d').drawImage(cv, 0, Math.round(sy), cv.width, Math.round(sh), 0, 0, cv.width, Math.round(sh));
      pdf.addImage(c2.toDataURL('image/jpeg', 0.82), 'JPEG', M, y, W, sh / pxPerMm);
      y += sh / pxPerMm;
    };
    const newPage = () => { pdf.addPage(); pages++; y = M; };
    // 続けて入る段落はまとめて1枚の画像にする
    let run = null;
    const flush = () => { if (run) { place(run.y, run.h); run = null; } };
    for (let i = 0; i < units.length; i++) {
      const u = units[i];
      if (u.pb) { flush(); if (y > M) newPage(); continue; }
      const next = units[i + 1];
      const gap = next && !next.pb ? Math.max(0, next.y - (u.y + u.h)) : 0;
      const hmm = (u.h + gap) / pxPerMm;
      const used = run ? run.h / pxPerMm : 0;
      if (y + used + hmm <= M + H) {
        if (!run) run = { y: u.y, h: 0 };
        run.h = u.y + u.h + gap - run.y;
        continue;
      }
      flush();
      if (hmm <= H) { if (y > M) newPage(); run = { y: u.y, h: u.h + gap }; continue; }
      // 1ページに入らない長い段落は切って並べる
      let off = u.y, rest = u.h;
      while (rest > 0) {
        if (y > M + H - 20) newPage();
        const part = Math.min(rest, (M + H - y) * pxPerMm);
        place(off, part);
        off += part; rest -= part;
        if (rest > 0) newPage();
      }
    }
    flush();
    for (let i = 1; i <= pages; i++) {
      pdf.setPage(i);
      pdf.setFontSize(8); pdf.setTextColor(150);
      pdf.text(`- ${i} / ${pages} -`, 105, 292, { align: 'center' });
    }
    stage.innerHTML = '';
    return pdf;
  }
  const baseName = (f) => f.replace(/\.docx$/i, '');
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  // ---------- 画面 ----------
  let results = [];
  let current = null;

  function showOne(res, fromList) {
    current = res;
    $('#oneView').classList.remove('hidden');
    $('#backBtn').classList.toggle('hidden', !fromList);
    $('#report').innerHTML = reportHtml(res);
    if (fromList) $('#oneView').scrollIntoView({ behavior: 'smooth' });
  }

  function showList() {
    $('#listView').classList.remove('hidden');
    $('#listCount').textContent = `${results.length}人分`;
    let h = '<tr><th>氏名・ファイル</th><th>ページ</th><th>アブスト</th><th>✕</th><th>△</th><th>まず直すこと</th></tr>';
    results.forEach((r, i) => {
      if (r.error) { h += `<tr><td>${esc(r.fileName)}</td><td colspan="5" style="color:var(--red)">${esc(r.error)}</td></tr>`; return; }
      const sm = summary(r);
      const s = r.stats;
      const pg = s.pagesWord ? s.pagesWord : '約' + s.pages;
      const pgBad = r.checks.pages.status !== 'ok';
      const abBad = r.checks.abstract.status === 'ng';
      h += `<tr class="lrow" data-i="${i}"><td><b>${esc(s.name || '')}</b><div class="muted">${esc(r.doc.fileName)}</div></td><td style="${pgBad ? 'color:var(--red);font-weight:700' : ''}">${pg}</td><td style="${abBad ? 'color:var(--red);font-weight:700' : ''}">${s.abstractLen}字</td><td><span class="pill ${sm.counts.ng ? 'ng' : 'ok'}">${sm.counts.ng}</span></td><td><span class="pill ${sm.counts.warn ? 'warn' : 'ok'}">${sm.counts.warn}</span></td><td>${sm.top.map((t) => esc(LABEL[t.key])).join('・') || '<span style="color:var(--ok)">なし</span>'}</td></tr>`;
    });
    $('#list').innerHTML = h;
    $('#list').querySelectorAll("tr.lrow").forEach((tr) => tr.addEventListener('click', () => showOne(results[+tr.dataset.i], true)));
  }

  async function handle(files) {
    files = Array.from(files).filter((f) => /\.docx$/i.test(f.name) && !f.name.startsWith('~$'));
    $('#msg').innerHTML = '';
    const skipped = files.length === 0;
    if (skipped) { $('#msg').innerHTML = '<div class="err">.docx のファイルを入れてください（古い .doc の場合は、Wordで「名前を付けて保存」→「Word文書（.docx）」にしてから）。</div>'; return; }
    results = [];
    for (const f of files) {
      try {
        const d = await Akapen.readDocx(await f.arrayBuffer(), f.name);
        results.push(Akapen.analyze(d));
      } catch (e) {
        console.error(e);
        results.push({ error: '読めませんでした：' + (e.message || e), fileName: f.name });
      }
    }
    results.sort((a, b) => ((a.stats && a.stats.name) || a.fileName || '').localeCompare((b.stats && b.stats.name) || b.fileName || '', 'ja', { numeric: true }));
    $('#listView').classList.add('hidden');
    $('#oneView').classList.add('hidden');
    if (results.length === 1) {
      if (results[0].error) { $('#msg').innerHTML = `<div class="err">${esc(results[0].error)}</div>`; return; }
      showOne(results[0], false);
      $('#oneView').scrollIntoView({ behavior: 'smooth' });
    } else {
      showList();
      $('#listView').scrollIntoView({ behavior: 'smooth' });
    }
  }

  // 入力
  const drop = $('#drop');
  $('#file').addEventListener('change', (e) => { handle(e.target.files); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => handle(e.dataTransfer.files));
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  $('#backBtn').addEventListener('click', () => $('#listView').scrollIntoView({ behavior: 'smooth' }));

  $('#pdfBtn').addEventListener('click', async () => {
    if (!current) return;
    const btn = $('#pdfBtn');
    btn.disabled = true; btn.textContent = 'PDFを作っています…';
    try {
      const pdf = await sheetToPdf(reportHtml(current));
      download(pdf.output('blob'), baseName(current.doc.fileName) + '_赤ペン.pdf');
    } catch (e) { console.error(e); alert('PDFを作れませんでした：' + e.message); }
    btn.disabled = false; btn.textContent = '赤ペンをPDFで保存';
  });

  $('#zipBtn').addEventListener('click', async () => {
    const ok = results.filter((r) => !r.error);
    if (!ok.length) return;
    const btn = $('#zipBtn');
    btn.disabled = true;
    $('#progBar').classList.remove('hidden');
    const zip = new JSZip();
    for (let i = 0; i < ok.length; i++) {
      $('#progTxt').textContent = `PDFを作っています ${i + 1} / ${ok.length}`;
      $('#prog').style.width = (i / ok.length * 100) + '%';
      const pdf = await sheetToPdf(reportHtml(ok[i]));
      zip.file(baseName(ok[i].doc.fileName) + '_赤ペン.pdf', pdf.output('arraybuffer'));
    }
    $('#prog').style.width = '100%';
    $('#progTxt').textContent = 'まとめています…';
    const blob = await zip.generateAsync({ type: 'blob' });
    download(blob, '卒論赤ペン_' + ok.length + '人分.zip');
    $('#progTxt').textContent = `${ok.length}人分を保存しました`;
    btn.disabled = false;
  });

  $('#csvBtn').addEventListener('click', () => {
    const head = ['氏名', 'ファイル', 'タイトル', 'ページ', 'アブストラクト字数', '✕の数', '△の数'].concat(CHECKS.map((c) => c[1]));
    const rows = [head];
    for (const r of results) {
      if (r.error) { rows.push(['', r.fileName, r.error]); continue; }
      const sm = summary(r);
      rows.push([r.stats.name || '', r.doc.fileName, r.stats.title || '', r.stats.pagesWord || ('約' + r.stats.pages), r.stats.abstractLen, sm.counts.ng, sm.counts.warn].concat(CHECKS.map((c) => MARK[(r.checks[c[0]] || {}).status] || '')));
    }
    const csv = '﻿' + rows.map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    download(new Blob([csv], { type: 'text/csv' }), '卒論チェック一覧.csv');
  });

  // 基準の一覧
  $('#rulesBody').innerHTML = '<table>' + CHECKS.map((c) => `<tr><td>${esc(c[1])}</td><td>${esc(c[2])}</td></tr>`).join('') + '</table><p class="muted">◯＝できている　△＝見直す　✕＝直す。ページ数は、Wordで最後に保存したときの数を使います。ファイルに正しい数が残っていないときは、字数からWindows版Wordでのページ数を換算します（±1ページずれることがあります）。内容のよしあしは判定しません。</p>';

  window.AkapenUI = { handle, reportHtml, sheetToPdf };
})();
