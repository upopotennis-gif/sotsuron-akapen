/* 卒論 赤ペンチェック：Wordファイル(.docx)を読んで体裁を判定するエンジン
   ファイルはブラウザの中だけで読み、どこにも送らない。 */
(function (global) {
  'use strict';

  // ---------------- 基準（フォーマット2026・論文ルールより） ----------------
  const RULE = {
    pageW: 11906, pageH: 16838,           // A4（twip）
    margin: 851,                          // 15mm
    marginBottomAlt: 567,                 // 見本の本文セクションは下10mm
    tol: 25,
    abstractMin: 350, abstractMax: 400,
    pagesMin: 4,
    expect: {
      title:   { sz: 24, b: true,  label: '明朝・太字・12ポイント' },
      name:    { sz: 18, b: false, label: '明朝・9ポイント' },
      abstract:{ sz: 18, b: false, label: '明朝・9ポイント' },
      keyword: { sz: 18, b: true,  label: '明朝・太字・9ポイント' },
      heading: { sz: 20, b: true,  label: '明朝・太字・10ポイント' },
      subheading: { sz: [18, 20], b: null, label: '明朝・9〜10ポイント' },
      refhead: { sz: 20, b: true,  label: '明朝・太字・10ポイント' },
      body:    { sz: 18, b: false, label: '明朝・9ポイント' },
      ref:     { sz: 18, b: false, label: '明朝・9ポイント' },
    },
  };

  const JA = /[぀-ヿ㐀-鿿ｦ-ﾟ]/;
  const tw2mm = (t) => Math.round(t / 56.7 * 10) / 10;
  const pt = (hp) => (hp / 2).toString().replace(/\.0$/, '');
  const norm = (s) => s.replace(/[\s　]/g, '');
  const countChars = (s) => [...norm(s)].length;
  const z2h = (s) => s.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));

  // ---------------- XMLの小道具 ----------------
  const kids = (el, tag) => el ? Array.from(el.children).filter((c) => !tag || c.tagName === tag) : [];
  const kid = (el, tag) => kids(el, tag)[0] || null;
  const attr = (el, name) => (el ? el.getAttribute(name) : null);
  const val = (el) => attr(el, 'w:val');
  const onOff = (el) => {
    if (!el) return undefined;
    const v = val(el);
    return !(v === '0' || v === 'false' || v === 'off');
  };
  function parseXml(str) {
    return new DOMParser().parseFromString(str, 'application/xml');
  }

  // ---------------- 書式（スタイル → 直接指定）の解決 ----------------
  function readRPr(rPr) {
    const o = {};
    if (!rPr) return o;
    const sz = kid(rPr, 'w:sz');
    if (sz) o.sz = +val(sz);
    const b = kid(rPr, 'w:b');
    if (b) o.b = onOff(b);
    const f = kid(rPr, 'w:rFonts');
    if (f) {
      if (attr(f, 'w:eastAsiaTheme')) { o.eaTheme = attr(f, 'w:eastAsiaTheme'); o.ea = null; }
      else if (attr(f, 'w:eastAsia')) { o.ea = attr(f, 'w:eastAsia'); o.eaTheme = null; }
      if (attr(f, 'w:asciiTheme')) o.asciiTheme = attr(f, 'w:asciiTheme');
      else if (attr(f, 'w:ascii')) o.ascii = attr(f, 'w:ascii');
    }
    const c = kid(rPr, 'w:color');
    if (c) o.color = val(c);
    const h = kid(rPr, 'w:highlight');
    if (h) o.highlight = val(h);
    const s = kid(rPr, 'w:shd');
    if (s && attr(s, 'w:fill') && !/^(auto|FFFFFF)$/i.test(attr(s, 'w:fill'))) o.shade = attr(s, 'w:fill');
    const va = kid(rPr, 'w:vertAlign');
    if (va) o.vert = val(va);
    const rs = kid(rPr, 'w:rStyle');
    if (rs) o.rStyle = val(rs);
    return o;
  }
  function merge(a, b) {
    const o = Object.assign({}, a);
    for (const k in b) if (b[k] !== undefined) o[k] = b[k];
    return o;
  }
  function readInd(pPr) {
    const ind = kid(pPr, 'w:ind');
    if (!ind) return null;
    const o = {};
    for (const a of ['firstLine', 'firstLineChars', 'hanging', 'hangingChars', 'left', 'leftChars']) {
      const v = attr(ind, 'w:' + a);
      if (v !== null) o[a] = +v;
    }
    return o;
  }

  function buildStyles(stylesDoc, themeDoc) {
    const theme = { minorEastAsia: '', majorEastAsia: '', minorHAnsi: '', majorHAnsi: '' };
    if (themeDoc) {
      for (const kind of ['minor', 'major']) {
        const f = themeDoc.getElementsByTagName('a:' + kind + 'Font')[0];
        if (!f) continue;
        const ea = kid(f, 'a:ea');
        let jp = '';
        for (const x of kids(f, 'a:font')) if (attr(x, 'script') === 'Jpan') jp = attr(x, 'typeface');
        theme[kind + 'EastAsia'] = (ea && attr(ea, 'typeface')) || jp || '';
        const la = kid(f, 'a:latin');
        theme[kind + 'HAnsi'] = la ? attr(la, 'typeface') : '';
      }
      // テーマに日本語フォントがない場合は Jpan を使う
      for (const kind of ['minor', 'major']) {
        const f = themeDoc.getElementsByTagName('a:' + kind + 'Font')[0];
        if (!f) continue;
        for (const x of kids(f, 'a:font')) if (attr(x, 'script') === 'Jpan' && !theme[kind + 'EastAsia']) theme[kind + 'EastAsia'] = attr(x, 'typeface');
      }
    }
    const map = {};
    let defR = {}, defP = {}, defaultPara = null;
    if (stylesDoc) {
      const root = stylesDoc.documentElement;
      const dd = kid(root, 'w:docDefaults');
      if (dd) {
        defR = readRPr(kid(kid(dd, 'w:rPrDefault'), 'w:rPr'));
        const pd = kid(kid(dd, 'w:pPrDefault'), 'w:pPr');
        if (pd) defP = { ind: readInd(pd) };
      }
      for (const s of kids(root, 'w:style')) {
        const id = attr(s, 'w:styleId');
        const pPr = kid(s, 'w:pPr');
        map[id] = {
          type: attr(s, 'w:type'),
          name: val(kid(s, 'w:name')) || id,
          basedOn: val(kid(s, 'w:basedOn')),
          r: readRPr(kid(s, 'w:rPr')),
          ind: pPr ? readInd(pPr) : null,
          numbered: !!(pPr && kid(pPr, 'w:numPr')),
          jc: pPr ? val(kid(pPr, 'w:jc')) : null,
        };
        if (attr(s, 'w:type') === 'paragraph' && /^(1|true)$/.test(attr(s, 'w:default') || '')) defaultPara = id;
      }
    }
    if (!defR.sz) defR.sz = 21; // Word の既定は10.5pt
    function chain(id, seen) {
      seen = seen || {};
      const s = map[id];
      if (!s || seen[id]) return [];
      seen[id] = 1;
      return chain(s.basedOn, seen).concat([s]);
    }
    function styleR(id) {
      return chain(id).reduce((acc, s) => merge(acc, s.r), {});
    }
    function styleP(id) {
      let ind = defP.ind ? Object.assign({}, defP.ind) : {};
      let numbered = false, jc = null;
      for (const s of chain(id)) {
        if (s.ind) ind = Object.assign(ind, s.ind);
        if (s.numbered) numbered = true;
        if (s.jc) jc = s.jc;
      }
      return { ind, numbered, jc };
    }
    function fontName(r) {
      if (r.eaTheme) return theme[r.eaTheme] || theme[r.eaTheme.replace(/Bidi$/, 'EastAsia')] || '';
      return r.ea || '';
    }
    return { map, defR, defaultPara, styleR, styleP, fontName, theme };
  }

  // ---------------- 段落・表を順に読む ----------------
  function readParagraph(p, st, ctx) {
    const pPr = kid(p, 'w:pPr');
    const pStyle = pPr ? val(kid(pPr, 'w:pStyle')) : null;
    const sid = pStyle || st.defaultPara;
    const sp = st.styleP(sid);
    const ind = Object.assign({}, sp.ind, pPr ? readInd(pPr) || {} : {});
    const baseR = merge(st.defR, st.styleR(sid));
    const blk = {
      type: 'p', text: '', runs: [], images: [], textboxes: [], footnotes: 0,
      styleId: sid, styleName: st.map[sid] ? st.map[sid].name : '',
      numbered: sp.numbered || !!(pPr && kid(pPr, 'w:numPr')),
      jc: (pPr && val(kid(pPr, 'w:jc'))) || sp.jc,
      ind, sectPr: pPr ? kid(pPr, 'w:sectPr') : null,
      pageBreakBefore: !!(pPr && kid(pPr, 'w:pageBreakBefore')),
    };
    function addText(t, r) {
      if (!t) return;
      const start = blk.text.length;
      blk.text += t;
      blk.runs.push({ start, end: blk.text.length, text: t, r });
    }
    function drawing(el) {
      // テキストボックスか画像か
      const tx = el.getElementsByTagName('w:txbxContent');
      if (tx.length) {
        for (const t of Array.from(tx)) blk.textboxes.push(Array.from(t.getElementsByTagName('w:t')).map((x) => x.textContent).join(''));
        return;
      }
      const ext = el.getElementsByTagName('wp:extent')[0];
      const anchor = el.getElementsByTagName('wp:anchor').length > 0;
      if (ext) blk.images.push({ cx: +attr(ext, 'cx'), cy: +attr(ext, 'cy'), anchor, at: blk.text.length });
      else if (el.getElementsByTagName('v:imagedata').length) blk.images.push({ cx: 0, cy: 0, anchor, at: blk.text.length });
    }
    function walkRun(run, inLink) {
      const rPr = kid(run, 'w:rPr');
      const direct = readRPr(rPr);
      let r = baseR;
      if (direct.rStyle) r = merge(r, st.styleR(direct.rStyle));
      r = merge(r, direct);
      if (inLink || (direct.rStyle && st.map[direct.rStyle] && /Hyperlink/i.test(st.map[direct.rStyle].name))) r = Object.assign({}, r, { link: true });
      for (const c of Array.from(run.children)) {
        switch (c.tagName) {
          case 'w:t': addText(c.textContent, r); break;
          case 'w:tab': addText('\t', r); break;
          case 'w:br': if (attr(c, 'w:type') === 'page') blk.pageBreak = true; else if (attr(c, 'w:type') !== 'column') addText('\n', r); break;
          case 'w:sym': addText('□', r); break;
          case 'w:footnoteReference': case 'w:endnoteReference': blk.footnotes++; addText('＊', Object.assign({}, r, { note: true })); break;
          case 'w:drawing': drawing(c); break;
          case 'w:pict': drawing(c); break;
          case 'mc:AlternateContent': {
            const ch = kid(c, 'mc:Choice') || kid(c, 'mc:Fallback');
            if (ch) for (const d of Array.from(ch.children)) drawing(d);
            break;
          }
        }
      }
    }
    function walk(el, inLink) {
      for (const c of Array.from(el.children)) {
        const t = c.tagName;
        if (t === 'w:r') walkRun(c, inLink);
        else if (t === 'w:del' || t === 'w:pPr' || t === 'w:moveFrom') continue;
        else if (t === 'mc:AlternateContent') {
          const ch = kid(c, 'mc:Choice');
          if (ch) walk(ch, inLink);
        } else if (t === 'w:hyperlink') walk(c, true);
        else if (['w:ins', 'w:smartTag', 'w:sdt', 'w:sdtContent', 'w:fldSimple', 'w:customXml', 'w:moveTo', 'w:dir', 'w:bdo'].includes(t)) walk(c, inLink);
      }
    }
    walk(p);
    return blk;
  }

  function readTable(tbl, st) {
    const rows = [], heights = [];
    let sz = 18;
    for (const tr of kids(tbl, 'w:tr')) {
      const cells = [];
      for (const tc of kids(tr, 'w:tc')) {
        const ps = kids(tc, 'w:p').map((p) => readParagraph(p, st));
        const r0 = ps.flatMap((p) => p.runs).find((r) => norm(r.text));
        if (r0 && r0.r.sz) sz = r0.r.sz;
        cells.push(ps.map((p) => p.text).join('\n'));
      }
      const trh = tr.getElementsByTagName('w:trHeight')[0];
      heights.push(trh ? +attr(trh, 'w:val') : 0);
      rows.push(cells);
    }
    const gridW = Array.from(tbl.getElementsByTagName('w:gridCol')).reduce((a, g) => a + (+attr(g, 'w:w') || 0), 0);
    return { type: 'table', rows, heights, gridW, sz, text: '', runs: [], images: [], textboxes: [], footnotes: 0, ind: {} };
  }

  function readSect(sp) {
    if (!sp) return null;
    const sz = kid(sp, 'w:pgSz'), mar = kid(sp, 'w:pgMar'), cols = kid(sp, 'w:cols'), grid = kid(sp, 'w:docGrid');
    return {
      w: sz ? +attr(sz, 'w:w') : null, h: sz ? +attr(sz, 'w:h') : null, orient: sz ? attr(sz, 'w:orient') : null,
      mar: mar ? { t: +attr(mar, 'w:top'), r: +attr(mar, 'w:right'), b: +attr(mar, 'w:bottom'), l: +attr(mar, 'w:left'), g: +(attr(mar, 'w:gutter') || 0) } : null,
      cols: cols && attr(cols, 'w:num') ? +attr(cols, 'w:num') : 1,
      colSpace: cols && attr(cols, 'w:space') ? +attr(cols, 'w:space') : 425,
      linePitch: grid && attr(grid, 'w:linePitch') ? +attr(grid, 'w:linePitch') : 360,
      type: val(kid(sp, 'w:type')) || 'nextPage',
    };
  }

  // ---------------- 読み込み本体 ----------------
  async function readDocx(arrayBuffer, fileName) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const get = async (p) => (zip.file(p) ? zip.file(p).async('string') : null);
    const docXml = await get('word/document.xml');
    if (!docXml) throw new Error('Wordファイル（.docx）として読めませんでした');
    const doc = parseXml(docXml);
    const st = buildStyles(parseXml((await get('word/styles.xml')) || '<x/>'), (await get('word/theme/theme1.xml')) ? parseXml(await get('word/theme/theme1.xml')) : null);
    const app = await get('docProps/app.xml');
    let pagesWord = null, application = '';
    if (app) {
      const m = app.match(/<Pages>(\d+)<\/Pages>/);
      if (m) pagesWord = +m[1];
      const a = app.match(/<Application>([^<]*)<\/Application>/);
      if (a) application = a[1];
    }
    const body = doc.getElementsByTagName('w:body')[0];
    const blocks = [];
    const sections = [];
    let secStart = 0;
    function pushBlocks(el) {
      for (const c of Array.from(el.children)) {
        if (c.tagName === 'w:p') {
          const b = readParagraph(c, st);
          b.idx = blocks.length;
          blocks.push(b);
          if (b.sectPr) {
            sections.push(Object.assign(readSect(b.sectPr), { from: secStart, to: b.idx }));
            secStart = b.idx + 1;
          }
        } else if (c.tagName === 'w:tbl') {
          const t = readTable(c, st);
          t.idx = blocks.length;
          blocks.push(t);
        } else if (c.tagName === 'w:sdt') {
          const sc = kid(c, 'w:sdtContent');
          if (sc) pushBlocks(sc);
        } else if (c.tagName === 'w:customXml') pushBlocks(c);
      }
    }
    pushBlocks(body);
    const last = kid(body, 'w:sectPr');
    sections.push(Object.assign(readSect(last) || { w: null, h: null, mar: null, cols: 1, colSpace: 425, linePitch: 360 }, { from: secStart, to: blocks.length - 1 }));
    blocks.forEach((b) => { b.sec = sections.findIndex((s) => b.idx >= s.from && b.idx <= s.to); });
    return { fileName, blocks, sections, pagesWord, application, st };
  }

  // ---------------- 役割（タイトル・アブストラクト・見出し…）を決める ----------------
  const RE_NAME = /^[\s　]*[0-9０-９]\s*[A-Za-zＡ-Ｚａ-ｚ]\s*[\(（]\s*[0-9０-９]{1,2}\s*[\)）]/;
  const RE_KEYWORD = /^[\s　]*(キーワード|キー・ワード|Key\s*words?)/i;
  const RE_REFHEAD = /^(参考|引用|引用・参考|参考・引用)?文献(一覧|リスト)?$|^参考資料$|^出典$/;
  const RE_CAPTION = /^[\s　]*[（(【\[]?\s*(図|表|グラフ|写真|資料|Fig\.?|Table)\s*[-－‐–—ー]?\s*[0-9０-９]/i;
  // 箇条書き・手順など、1字下げや句点を求めない行
  const RE_LISTITEM = /^[\s　]*([・●○◆◇■□※★☆▶→⇨⇒〇◎◯-]|[①-⑳]|[\(（][0-9０-９]{1,2}[\)）]|[\(（][一二三四五六七八九十]{1,3}[\)）]|[0-9０-９]{1,2}[\)）．.、]|手順|第[一二三四五六七八九十0-9０-９]+[回段章項]|[ア-オ][\)）．.、]|[a-eA-E][\)）．.])/;
  const RE_SOURCE = /^[\s　]*[（(【]?\s*(出典|出所|出展|資料|引用元|データ)\s*[)）】:：]/;
  const RE_HEADNUM = /^[\s　]*([0-9０-９]{1,2})(\s*[\.．、）)]|[-‐－][0-9０-９])|^[\s　]*第[一二三四五六七八九十0-9０-９]+章|^[\s　]*[■●◆□◇]/;

  function assignRoles(d) {
    const B = d.blocks;
    const nonEmpty = (b) => b.type === 'table' || norm(b.text) !== '' || b.images.length > 0;
    // 見出し部分（タイトル〜キーワード）の範囲
    let headerEnd = -1;
    const twoCol = d.sections.findIndex((s) => s.cols >= 2);
    const kwIdx = B.findIndex((b) => b.type === 'p' && RE_KEYWORD.test(b.text));
    if (twoCol > 0) headerEnd = d.sections[twoCol - 1].to;
    else if (kwIdx >= 0) headerEnd = kwIdx;
    else {
      const h = B.findIndex((b, i) => i > 0 && b.type === 'p' && RE_HEADNUM.test(b.text) && norm(b.text).length < 45);
      headerEnd = h > 0 ? h - 1 : Math.min(3, B.length - 1);
    }
    if (kwIdx > headerEnd) headerEnd = kwIdx;
    d.headerEnd = headerEnd;
    let stage = 'title';
    for (const b of B) {
      if (b.idx > headerEnd) break;
      if (!nonEmpty(b)) { b.role = 'empty'; continue; }
      if (b.type === 'table') { b.role = 'table'; continue; }
      if (RE_KEYWORD.test(b.text)) { b.role = 'keyword'; stage = 'after'; continue; }
      if (stage === 'title') { b.role = 'title'; stage = 'subtitle'; continue; }
      if (stage === 'subtitle') {
        if (RE_NAME.test(b.text)) { b.role = 'name'; stage = 'abstract'; continue; }
        if (/^[\s　]*[―—─‐\-〜～~－ー–]/.test(b.text) || (!RE_NAME.test(b.text) && B.some((x) => x.idx > b.idx && x.idx <= headerEnd && RE_NAME.test(x.text)))) { b.role = 'subtitle'; continue; }
        if (norm(b.text).length < 40) { b.role = 'name'; stage = 'abstract'; continue; }
        stage = 'abstract';
      }
      if (stage === 'abstract') { b.role = 'abstract'; continue; }
      b.role = 'other';
    }
    // 本文
    let inRefs = false;
    for (const b of B) {
      if (b.idx <= headerEnd) continue;
      if (b.type === 'table') { b.role = 'table'; continue; }
      if (!nonEmpty(b)) { b.role = 'empty'; continue; }
      const t = norm(b.text);
      if (!t && b.images.length) { b.role = 'figure'; continue; }
      if (RE_REFHEAD.test(t)) { b.role = 'refhead'; inRefs = true; continue; }
      if (inRefs) { b.role = 'ref'; continue; }
      if (RE_CAPTION.test(b.text) && t.length < 90) { b.role = 'caption'; continue; }
      if (RE_SOURCE.test(b.text) && t.length < 120) { b.role = 'source'; continue; }
      const allBold = b.runs.filter((r) => norm(r.text)).every((r) => r.r.b);
      const endsSentence = /[。．.」]$/.test(t);
      if (t.length <= 45 && !endsSentence && (RE_HEADNUM.test(b.text) || b.numbered || allBold)) { b.role = 'heading'; continue; }
      b.role = 'body';
      if (b.images.length) b.hasFigure = true;
    }
    // 章の見出しと節の見出し（章の中の「1．」「1-1」など）を分ける
    const numOf = (b) => { const m = z2h(b.text).match(/^[\s　]*(\d{1,2})(?![-‐－.．]\d)/); return m ? +m[1] : null; };
    const sigOf = (b) => { const r = b.runs.find((x) => norm(x.text)); return r ? `${r.r.sz}/${!!r.r.b}` : ''; };
    let lastCh = 0, chSig = null, inSub = false, lastSub = 0, subSig = null;
    const RE_CHAPWORD = /問題提起|まえがき|はじめに|序論|まとめ|おわりに|終わりに|結論|結び|むすび|今後の課題/;
    for (const b of B) {
      if (b.role !== 'heading') continue;
      const n = numOf(b), sig = sigOf(b);
      const indented = /^[ 　]{1,}/.test(b.text);
      const dotted = /^[\s　]*[0-9０-９]+\s*[-‐－.．][0-9０-９]+/.test(b.text);
      let sub = false;
      if (chSig !== null) {
        if (indented || dotted) sub = true;
        else if (RE_CHAPWORD.test(b.text)) sub = false;
        else if (inSub && n !== null && n === lastSub + 1 && sig === subSig && sig !== chSig) sub = true;
        else if (n !== null && n === lastCh + 1) sub = false;
        else if (n !== null && (n < lastCh || (n === lastCh && sig !== chSig))) sub = true;
        else if (inSub && sig !== chSig) sub = true;
      }
      if (sub) { b.role = 'subheading'; inSub = true; lastSub = n !== null ? n : lastSub + 1; subSig = sig; continue; }
      if (chSig === null) chSig = sig;
      lastCh = n !== null ? n : lastCh + 1;
      inSub = false; lastSub = 0;
    }
    return d;
  }

  // ---------------- 判定 ----------------
  function isMincho(f) { return /明朝|Mincho|Minch|Serif|明體|宋体|Song|Shippori|Zen Old|Hina|HiraMin|Kozuka|Ryumin|リュウミン|教科書|Kyokasho/i.test(f); }
  function isGothic(f) { return /ゴシック|Gothic|Meiryo|メイリオ|Sans|Kaku|角ゴ|HiraKaku|Arial|Helvetica|Calibri|丸ゴ|Maru|Osaka|ポップ|Pop|行書|楷書|Kaisho|Gyosho|HG/i.test(f); }

  function analyze(d) {
    assignRoles(d);
    const B = d.blocks;
    const issues = [];   // {level, key, block, spans:[[s,e]], msg, fix}
    const add = (o) => { issues.push(o); return o; };
    const byRole = (r) => B.filter((b) => b.role === r);
    const checks = {};   // key -> {status, detail}
    const setCheck = (key, status, detail) => {
      const order = { ok: 0, hint: 0, warn: 1, ng: 2 };
      if (!checks[key] || order[status] > order[checks[key].status]) checks[key] = { status, detail: detail || '' };
      else if (detail && order[status] === order[checks[key].status] && !checks[key].detail) checks[key].detail = detail;
    };
    const stats = {};

    // ---- 1 用紙・余白 ----
    {
      const probs = [];
      d.sections.forEach((s, i) => {
        if (!s.w) return;
        if (s.orient === 'landscape' || s.w > s.h) probs.push('用紙が横向きになっています → 縦にしよう');
        else if (Math.abs(s.w - RULE.pageW) > 60 || Math.abs(s.h - RULE.pageH) > 60) probs.push(`用紙が${tw2mm(s.w)}×${tw2mm(s.h)}mmです → A4（210×297mm）にしよう`);
        if (s.mar) {
          const side = { t: '上', r: '右', b: '下', l: '左' };
          for (const k of ['t', 'r', 'b', 'l']) {
            const v = s.mar[k];
            const ok = Math.abs(v - RULE.margin) <= RULE.tol || (k === 'b' && Math.abs(v - RULE.marginBottomAlt) <= RULE.tol);
            if (!ok) probs.push(`余白（${side[k]}）が${tw2mm(v)}mmです → 15mmにしよう`);
          }
          if (s.mar.g > 0) probs.push(`とじしろが${tw2mm(s.mar.g)}mmあります → 0mmにしよう`);
        }
      });
      const uniq = [...new Set(probs)];
      if (uniq.length) {
        setCheck('paper', 'ng', uniq.join('／'));
        uniq.forEach((m) => add({ level: 'ng', key: 'paper', block: null, msg: m.split(' → ')[0], fix: m.split(' → ')[1] || '' }));
      } else setCheck('paper', 'ok', 'A4縦・余白15mm');
    }

    // ---- 2 段組み ----
    {
      const bodyBlocks = B.filter((b) => b.idx > d.headerEnd && b.role !== 'empty');
      const bodySecs = [...new Set(bodyBlocks.map((b) => b.sec))];
      const oneColBody = bodySecs.filter((i) => d.sections[i] && d.sections[i].cols < 2);
      const headSecs = [...new Set(B.filter((b) => b.idx <= d.headerEnd && b.role !== 'empty').map((b) => b.sec))];
      if (!bodySecs.length) setCheck('columns', 'ng', '本文が見つかりません');
      else if (oneColBody.length === bodySecs.length) {
        setCheck('columns', 'ng', '本文が1段のままです');
        add({ level: 'ng', key: 'columns', block: bodyBlocks[0] ? bodyBlocks[0].idx : null, msg: '本文が2段組みになっていません', fix: 'キーワードの後から最後までを選んで［レイアウト］→［段組み］→［2段］' });
      } else if (oneColBody.length) {
        setCheck('columns', 'warn', '本文の一部が1段になっています');
        const b = bodyBlocks.find((x) => oneColBody.includes(x.sec));
        add({ level: 'warn', key: 'columns', block: b ? b.idx : null, msg: 'ここから先が1段組みになっています', fix: '本文はすべて2段組みにそろえよう' });
      } else setCheck('columns', 'ok', '本文2段組み');
      if (headSecs.some((i) => d.sections[i] && d.sections[i].cols >= 2)) {
        setCheck('columns', 'warn', 'タイトル〜キーワードも2段になっています');
        add({ level: 'warn', key: 'columns', block: 0, msg: 'タイトル〜キーワードまで2段組みになっています', fix: 'ここは1段、本文から2段にしよう（キーワードの後に「セクション区切り（現在の位置から）」）' });
      }
    }

    // ---- 3 ページ数 ----
    {
      const est = estimatePages(d);
      stats.pagesEst = est;
      // Word が保存した値は、Web版・自動保存などで「1」のまま残ることがあるので、字数の目安と大きく食い違うときは使わない
      const pw = d.pagesWord;
      const fromWord = pw && /Word/i.test(d.application) && pw > 1 && pw >= est * 0.6 && pw <= est * 1.6 + 1 ? pw : null;
      stats.pagesWord = fromWord;
      stats.pages = fromWord || Math.round(est);
      let status = 'ok', fix = '', how;
      if (fromWord) {
        const n = fromWord;
        how = `Wordで最後に保存したときのページ数：${n}ページ`;
        if (n < RULE.pagesMin) { status = 'ng'; fix = `最低4ページ必要です。あと${RULE.pagesMin - n}ページ分書き足そう`; }
        else if (n % 2 === 1) { status = 'ng'; fix = `製本のため偶数ページにします。${n - 1}ページに収めるか、${n + 1}ページまで書き足そう`; }
      } else {
        // Windows版Wordで開いたときのページ数の目安（最後のページの途中まででも1ページ）
        const n = Math.max(1, Math.ceil(est - 0.02));
        const edge = est - Math.floor(est) < 0.15 || Math.ceil(est) - est < 0.1;
        stats.pages = n;
        how = `Windows版Wordでの目安：約${n}ページ${edge ? '（' + (est - Math.floor(est) < 0.15 ? `${n - 1}〜${n}` : `${n}〜${n + 1}`) + 'ページの境目）' : ''}。ファイルに正しいページ数が残っていないため字数から換算`;
        if (est < 2.5) { status = 'ng'; fix = '最低4ページ必要です。書き足そう'; }
        else if (n < RULE.pagesMin || n % 2 === 1 || edge) { status = 'warn'; fix = '学校のPCのWordで、画面左下のページ数を確かめよう（4ページ以上の偶数）'; }
      }
      setCheck('pages', status, how + (fix ? '／' + fix : ''));
      if (status !== 'ok') add({ level: status, key: 'pages', block: null, msg: fromWord ? `${fromWord}ページになっています` : `Windows版Wordで約${stats.pages}ページの見込みです（目安）`, fix });
    }

    // ---- 4 見本の文字が残っていないか ----
    {
      const LEFT = [
        /明朝\s*(太\s*)?[0-9０-９]+\s*ポイント/g, /学年ルーム\(番号\)氏名/g, /山田たろう/g, /オートクチュール/g, /ビスポーク/g,
        /どこへ越しても住みにくい/g, /〜ヘンおよび標準語形/g, /[●■▲◇]{2}/g, /久保光志/g, /木坂基/g, /中村明（1974）/g,
        /（問題提起を受けて）/g, /（主張が正しいことの証明[^）]*）/g, /よりづか★ちょいスポ/g, /智に働けば角が立つ/g,
      ];
      let found = 0;
      for (const b of B) {
        if (b.type !== 'p') continue;
        const spans = [];
        for (const re of LEFT) { re.lastIndex = 0; let m; while ((m = re.exec(b.text))) spans.push([m.index, m.index + m[0].length]); }
        if (spans.length) {
          b.leftSpans = spans;
          found++;
          add({ level: 'ng', key: 'leftover', block: b.idx, spans, msg: '見本（フォーマット）の文字が残っています', fix: '自分の文に書きかえるか、消そう' });
        }
        for (const tb of b.textboxes) {
          if (/A4縦|余白\s*上下左右|段組み設定|最低\s*[4４]\s*ページ|アブストラクト\s*350|明朝\s*[0-9０-９]+\s*ポイント/.test(tb)) {
            found++;
            add({ level: 'ng', key: 'leftover', block: b.idx, msg: `見本の説明の枠「${tb.slice(0, 22)}${tb.length > 22 ? '…' : ''}」が残っています`, fix: '枠の線をクリックして Delete で消そう' });
          }
        }
      }
      // 同じ説明の枠が2つずつ入っていることがあるので重複をまとめる
      const seen = new Set();
      for (let i = issues.length - 1; i >= 0; i--) {
        const x = issues[i];
        if (x.key !== 'leftover' || x.spans) continue;
        const k = x.block + x.msg;
        if (seen.has(k)) issues.splice(i, 1); else seen.add(k);
      }
      setCheck('leftover', found ? 'ng' : 'ok', found ? `${issues.filter((x) => x.key === 'leftover').length}か所` : '残っていません');
    }

    // ---- 書式（大きさ・太さ・明朝）を役割ごとに ----
    // 見本の書き込み（「明朝9ポイント」など）は、残っていることを別に指摘するので書式の判定からは外す
    const inLeftover = (b, r) => (b.leftSpans || []).some(([s0, e0]) => r.start >= s0 - 1 && r.end <= e0 + 1);
    function fontCheck(b, role, key) {
      const ex = RULE.expect[role];
      if (!ex || b.type !== 'p') return;
      const runs = b.runs.filter((r) => norm(r.text) && !r.r.note && r.r.vert !== 'superscript' && r.r.vert !== 'subscript' && !inLeftover(b, r));
      if (!runs.length) return;
      const szOk = (v) => (Array.isArray(ex.sz) ? ex.sz.includes(v) : v === ex.sz);
      const szBad = runs.filter((r) => r.r.sz && !szOk(r.r.sz));
      const bBad = ex.b === null ? [] : runs.filter((r) => !!r.r.b !== ex.b);
      const fontBad = runs.filter((r) => JA.test(r.text)).map((r) => ({ r, f: d.st.fontName(r.r) })).filter((x) => x.f && !isMincho(x.f));
      const chars = (arr) => arr.reduce((n, r) => n + norm(r.text).length, 0);
      const total = chars(runs);
      const probs = [];
      let lvl = 'warn';
      if (szBad.length) {
        const sizes = [...new Set(szBad.map((r) => pt(r.r.sz)))];
        probs.push(`${sizes.join('・')}ポイントになっています`);
      }
      if (bBad.length && chars(bBad) >= Math.min(2, total)) probs.push(ex.b ? '太字になっていません' : '太字になっています');
      if (fontBad.length) {
        const f = [...new Set(fontBad.map((x) => x.f))];
        probs.push(`「${f.join('」「')}」になっています`);
        if (fontBad.some((x) => isGothic(x.f))) lvl = 'ng';
      }
      if (szBad.length && chars(szBad) / total > 0.5) lvl = 'ng';
      if (!probs.length) return;
      const spans = (szBad.length || fontBad.length) && chars(szBad.concat(fontBad.map((x) => x.r))) < total * 0.6
        ? szBad.concat(fontBad.map((x) => x.r)).map((r) => [r.start, r.end]) : null;
      add({ level: lvl, key, block: b.idx, spans, msg: probs.join('。'), fix: `${ex.label}にそろえよう` });
      return lvl;
    }

    // ---- 5 タイトル ----
    {
      const t = byRole('title')[0];
      if (!t) { setCheck('title', 'ng', 'タイトルが見つかりません'); add({ level: 'ng', key: 'title', block: 0, msg: 'タイトルが見つかりません', fix: '1行目に論文のタイトルを書こう' }); }
      else {
        const r = fontCheck(t, 'title', 'title');
        if (t.jc !== 'center') add({ level: 'warn', key: 'title', block: t.idx, msg: 'タイトルが中央にそろっていません', fix: '［ホーム］→［中央ぞろえ］' });
        setCheck('title', r || (t.jc !== 'center' ? 'warn' : 'ok'), r ? '書式を直そう' : (t.jc !== 'center' ? '中央ぞろえに' : '明朝・太字・12pt'));
        stats.title = norm(t.text);
      }
      byRole('subtitle').forEach((b) => fontCheck(b, 'title', 'title'));
    }

    // ---- 6 氏名 ----
    {
      const n = byRole('name')[0];
      if (!n) { setCheck('name', 'ng', '氏名の行が見つかりません'); add({ level: 'ng', key: 'name', block: byRole('title')[0] ? byRole('title')[0].idx : 0, msg: 'タイトルの下に氏名の行がありません', fix: '「3C(12)氏名」の形で書こう' }); }
      else {
        const r = fontCheck(n, 'name', 'name');
        const okForm = RE_NAME.test(n.text);
        if (!okForm) add({ level: 'warn', key: 'name', block: n.idx, msg: '氏名の書き方が見本と違います', fix: '「学年クラス(番号)氏名」＝「3C(12)川崎花子」の形で' });
        setCheck('name', r || (okForm ? 'ok' : 'warn'), okForm ? '学年ルーム(番号)氏名' : '書き方を見本に合わせよう');
        let nm = n.text;
        (n.leftSpans || []).slice().sort((x, y) => y[0] - x[0]).forEach(([s0, e0]) => { nm = nm.slice(0, s0) + nm.slice(e0); });
        stats.name = nm.replace(/[\s　]+/g, ' ').trim();
      }
    }

    // ---- 7 アブストラクト ----
    {
      const ab = byRole('abstract');
      const len = ab.reduce((s, b) => s + countChars(b.text), 0);
      stats.abstractLen = len;
      if (!ab.length || len === 0) {
        setCheck('abstract', 'ng', 'アブストラクトが見つかりません');
        add({ level: 'ng', key: 'abstract', block: byRole('name')[0] ? byRole('name')[0].idx : 0, msg: 'アブストラクト（要旨）が見つかりません', fix: '氏名の下に、論文全体の要約を350〜400字で書こう' });
      } else {
        let st = 'ok', fix = '';
        if (len < RULE.abstractMin) st = 'ng', fix = `あと${RULE.abstractMin - len}字以上書き足そう（350〜400字）`;
        else if (len > RULE.abstractMax) st = 'ng', fix = `${len - RULE.abstractMax}字以上けずろう（350〜400字）`;
        if (st !== 'ok') add({ level: 'ng', key: 'abstract', block: ab[0].idx, msg: `アブストラクトが${len}字です`, fix });
        let fr = null;
        ab.forEach((b) => { fr = fontCheck(b, 'abstract', 'abstract') || fr; });
        setCheck('abstract', st === 'ng' ? 'ng' : (fr || 'ok'), `${len}字（350〜400字）`);
      }
    }

    // ---- 8 キーワード ----
    {
      const k = byRole('keyword')[0];
      if (!k) {
        setCheck('keyword', 'ng', 'キーワードの行がありません');
        const ab = byRole('abstract');
        add({ level: 'ng', key: 'keyword', block: ab.length ? ab[ab.length - 1].idx : 0, msg: 'キーワードの行がありません', fix: 'アブストラクトの下に「キーワード：〇〇、〇〇、〇〇」を書こう' });
      } else {
        const kwBody = k.text.replace(RE_KEYWORD, '').replace(/^[\s　:：]+/, '').trim();
        const words = kwBody.split(/[、,，・\/／\s　]+/).map((s) => norm(s)).filter(Boolean);
        if (words.length > 1 && !/[、,，]/.test(kwBody)) add({ level: 'hint', key: 'keyword', block: k.idx, msg: 'キーワードの区切りが「、」になっていません', fix: '見本のように「、」で区切ろう' });
        stats.keywords = words;
        const r = fontCheck(k, 'keyword', 'keyword');
        if (!/キーワード：/.test(k.text)) add({ level: 'warn', key: 'keyword', block: k.idx, msg: '「キーワード：」の形になっていません', fix: '「キーワード：」（全角のコロン）のあとに語を並べよう' });
        if (words.length < 2) add({ level: 'warn', key: 'keyword', block: k.idx, msg: `キーワードが${words.length}語しかありません`, fix: '論文の中心になる語を「、」で区切っていくつか挙げよう' });
        setCheck('keyword', r || (words.length < 2 ? 'warn' : 'ok'), `${words.length}語`);
      }
    }

    // ---- 9 章立て ----
    {
      const hs = byRole('heading');
      stats.headings = hs.map((h) => norm(h.text));
      const refhead = byRole('refhead')[0];
      if (!hs.length) {
        setCheck('structure', 'ng', '見出しが見つかりません');
        const first = B.find((b) => b.role === 'body');
        add({ level: 'ng', key: 'structure', block: first ? first.idx : null, msg: '章の見出しが見つかりません', fix: '「1．問題提起」「2．〇〇」…のように、章ごとに番号つきの見出しを立てよう' });
      } else {
        let st = 'ok', detail = hs.length + '章';
        // 章の名前・順番は自由（本人の方針）。見出しがあるか・番号・文献の見出しだけを見る
        // 番号の飛び
        const nums = hs.map((h) => { const m = z2h(h.text).match(/^[\s　]*(\d{1,2})(?![-‐－\d])/); return m ? +m[1] : null; }).filter((x) => x !== null);
        for (let i = 1; i < nums.length; i++) if (nums[i] !== nums[i - 1] + 1 && nums[i] !== 1) {
          const h = hs.filter((x) => /^[\s　]*[0-9０-９]/.test(x.text))[i];
          if (h) add({ level: 'warn', key: 'structure', block: h.idx, msg: nums[i] === nums[i - 1] ? `章の番号「${nums[i]}」が2回あります` : `章の番号が${nums[i - 1]}の次が${nums[i]}になっています`, fix: '番号を順にそろえよう' });
          if (st === 'ok') st = 'warn', detail = '章の番号がそろっていません';
          break;
        }
        if (!refhead) { st = 'ng'; detail += '／「文献」の見出しがありません'; add({ level: 'ng', key: 'refs', block: hs[hs.length - 1].idx, msg: '「文　献」の見出しが見つかりません', fix: '最後に「文　献」の見出しを立て、使った本・論文・サイトを並べよう' }); }
        setCheck('structure', st, detail);
      }
    }

    // ---- 10 見出しの文字 ----
    {
      let worst = null;
      byRole('heading').concat(byRole('refhead'), byRole('subheading')).forEach((h) => {
        const r = fontCheck(h, h.role, 'headfont');
        if (r === 'ng' || (r && !worst)) worst = r;
      });
      setCheck('headfont', worst || 'ok', worst ? '見出しの書式を直そう' : '明朝・太字・10pt');
    }

    // ---- 11 本文の文字 ----
    {
      let worst = null, cnt = 0;
      byRole('body').forEach((b) => { const r = fontCheck(b, 'body', 'bodyfont'); if (r) { cnt++; if (r === 'ng' || !worst) worst = r; } });
      byRole('ref').forEach((b) => { const r = fontCheck(b, 'ref', 'bodyfont'); if (r) { cnt++; if (r === 'ng' || !worst) worst = r; } });
      setCheck('bodyfont', worst || 'ok', worst ? `${cnt}段落で書式が違います` : '明朝・9pt');
    }

    // ---- 12 段落の書き出し（1字下げ）・文の途中の改行 ----
    {
      const bodies = byRole('body');
      let noInd = 0, dbl = 0, broken = 0;
      bodies.forEach((b, i) => {
        const lead = b.text.match(/^[ \t　]*/)[0];
        const hasLeadSpace = /　|  /.test(lead);
        const hasIndent = (b.ind.firstLineChars || 0) > 0 || (b.ind.firstLine || 0) > 60;
        const startsQuote = /^[\s　]*[「『（(]/.test(b.text);
        // 箇条書き・左に寄せた引用（抜き出し型）・ぶら下げなどは字下げを求めない
        const special = RE_LISTITEM.test(b.text) || (b.ind.left || 0) > 60 || (b.ind.leftChars || 0) > 0 || (b.ind.hanging || 0) > 0 || (b.ind.firstLineChars || 0) >= 200;
        const listItem = special;
        if (!hasLeadSpace && !hasIndent && !startsQuote && !b.hasFigure && !listItem) {
          noInd++;
          add({ level: 'warn', key: 'indent', block: b.idx, spans: [[lead.length, Math.min(lead.length + 1, b.text.length)]], msg: '段落の書き出しが下がっていません', fix: '書き出しを1マスあけよう' });
        } else if (hasLeadSpace && hasIndent && (b.ind.firstLineChars || 0) >= 100) {
          dbl++;
          add({ level: 'hint', key: 'indent', block: b.idx, spans: [[0, lead.length]], msg: '書き出しが2マス分下がっています（字下げ設定＋全角スペース）', fix: 'どちらか一方にして1マスに' });
        }
        const t = b.text.replace(/[\s　]+$/, '');
        const next = B[b.idx + 1];
        if (t && t.length >= 25 && !listItem && !/[。．.」』）)！？!?：:；;]$/.test(t) && next && next.role === 'body' && !RE_LISTITEM.test(next.text)) {
          broken++;
          add({ level: 'warn', key: 'indent', block: b.idx, spans: [[t.length - 1, t.length]], msg: '文の途中で改行（Enter）しているようです', fix: '次の段落とつなげよう（行の折り返しはWordに任せる）' });
        }
      });
      const st = noInd || broken ? 'warn' : 'ok';
      setCheck('indent', st, st === 'ok' ? '1字下げできています' : [noInd ? `下がっていない段落が${noInd}` : '', broken ? `途中で改行が${broken}` : ''].filter(Boolean).join('／'));
    }

    // ---- 13 文体（です・ます） と 言葉づかい ----
    function quoteMask(text) {
      const mask = new Array(text.length).fill(false);
      let depth = 0;
      for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (c === '「' || c === '『') depth++;
        if (depth > 0) mask[i] = true;
        if ((c === '」' || c === '』') && depth > 0) depth--;
      }
      return mask;
    }
    {
      const RE_KEI = /(ですが|です|ます|でした|ました|ません|でしょう|ましょう|ください)(?=[。、，．,」』）)！？!?\s　]|$|か[。、]|ね|よ[。、]|が|し[、。]|の[でに])/g;
      let cnt = 0, paras = 0;
      const targets = B.filter((b) => ['body', 'abstract', 'heading', 'caption'].includes(b.role));
      for (const b of targets) {
        const mask = quoteMask(b.text);
        const spans = [];
        let m;
        RE_KEI.lastIndex = 0;
        while ((m = RE_KEI.exec(b.text))) {
          if (mask[m.index]) continue;
          if (b.text.slice(m.index, m.index + 4) === 'ますま' || b.text.slice(m.index - 2, m.index) === 'ます') continue; // 「ますます」
          spans.push([m.index, m.index + m[0].length]);
        }
        if (spans.length) {
          cnt += spans.length; paras++;
          add({ level: 'ng', key: 'style', block: b.idx, spans, msg: `「です・ます」になっています${spans.length > 1 ? `（この段落に${spans.length}か所）` : ''}`, fix: '「〜である」「〜だ」に直そう' });
        }
      }
      setCheck('style', cnt ? 'ng' : 'ok', cnt ? `${paras}段落・${cnt}か所` : '「である」調');
      stats.keitai = cnt;

      // 話し言葉など（ルールの外の、もう一歩の助言）
      const TALK = [
        [/じゃ(ない|なく|なかっ)/g, '「じゃ」は話し言葉', '「では」に'],
        [/(?<![おわ])けど(?!う)/g, '「けど」は話し言葉', '「が」「けれども」に'],
        [/すご(く|い)/g, '「すごい」は話し言葉', '「非常に」「大きく」など、程度を具体的に'],
        [/やっぱり|やっぱ/g, '「やっぱり」は話し言葉', '「やはり」に'],
        [/みたい(な|に|だ)/g, '「みたい」は話し言葉', '「のような」「のように」に'],
        [/(?<=[ぁ-んァ-ヶ一-龠])とか(?![らん])/g, '「とか」は話し言葉', '「や」「など」に'],
        [/ちゃんと/g, '「ちゃんと」は話し言葉', '「きちんと」「適切に」に'],
        [/いっぱい/g, '「いっぱい」は話し言葉', '「多く」「多数」に'],
        [/(^|[。\s　])(なので|だから|でも|だけど)(?=[、，])/g, '文の始めの「なので」「だから」「でも」は話し言葉', '「そのため」「したがって」「しかし」に'],
        [/[!！]/g, '論文では「！」を使わない', '「。」にしよう'],
        [/(?<![A-Za-z])(ok|OK|Ok)(?![A-Za-z])/g, '「ok」は話し言葉', '「許可」「了承」など'],
      ];
      const OMOU = /と(思う|思った|思われる|思います|思いました|感じる|感じた)/g;
      let omou = 0;
      const omouHits = [];
      for (const b of B.filter((x) => ['body', 'abstract'].includes(x.role))) {
        const mask = quoteMask(b.text);
        for (const [re, msg, fix] of TALK) {
          re.lastIndex = 0;
          const spans = [];
          let m;
          while ((m = re.exec(b.text))) {
            const off = m[1] && re.source.startsWith('(^') ? m[1].length : 0;
            if (mask[m.index + off]) continue;
            spans.push([m.index + off, m.index + m[0].length]);
            if (m[0].length === 0) re.lastIndex++;
          }
          if (spans.length) add({ level: 'hint', key: 'talk', block: b.idx, spans, msg, fix });
        }
        OMOU.lastIndex = 0;
        const sp = [];
        let m;
        while ((m = OMOU.exec(b.text))) if (!mask[m.index]) sp.push([m.index, m.index + m[0].length]);
        if (sp.length) { omou += sp.length; omouHits.push({ b, sp }); }
      }
      if (omou >= 3) omouHits.forEach(({ b, sp }) => add({ level: 'hint', key: 'talk', block: b.idx, spans: sp, msg: '「〜と思う」が多めです（論文全体で' + omou + 'か所）', fix: '根拠を示して「〜と考える」「〜である」と言い切ろう' }));
      stats.omou = omou;
    }

    // ---- 14 引用の書き方 ----
    {
      let n = 0;
      for (const b of B.filter((x) => ['body', 'abstract'].includes(x.role))) {
        const sp = [];
        let m;
        const re = /[。．]」/g;
        while ((m = re.exec(b.text))) sp.push([m.index, m.index + 1]);
        if (sp.length) { n++; add({ level: 'warn', key: 'quote', block: b.idx, spans: sp, msg: '「」の最後に「。」が入っています', fix: '「」の最後の「。」は取ろう（原文の句点まで引用する場合を除く）' }); }
        // 「」の中の「」
        const sp2 = [];
        let depth = 0;
        for (let i = 0; i < b.text.length; i++) {
          if (b.text[i] === '「') { if (depth >= 1) sp2.push([i, i + 1]); depth++; }
          if (b.text[i] === '」' && depth > 0) { depth--; if (depth >= 1) sp2.push([i, i + 1]); }
        }
        if (sp2.length) { n++; add({ level: 'warn', key: 'quote', block: b.idx, spans: sp2, msg: '「」の中に「」があります', fix: '中の「」は『』に直そう' }); }
      }
      setCheck('quote', n ? 'warn' : 'ok', n ? `${n}か所` : '問題なし');
    }

    // ---- 15 文献 ----
    {
      // 1つの文献が2行（文献＋URLなど）に分かれていることが多いので、まとめて1件にする
      const refBlocks = byRole('ref').filter((b) => norm(b.text));
      const NUM = /^[\s　]*[\(（]\s*[0-9０-９]{1,2}\s*[\)）]/;
      const numberedStyle = refBlocks.filter((b) => NUM.test(b.text)).length >= refBlocks.length / 2;
      const refs = [];
      for (const b of refBlocks) {
        const cont = /^[\s　]*[\(（]?\s*(https?:|www\.|file:)/i.test(b.text) || (numberedStyle && !NUM.test(b.text));
        if (cont && refs.length) { refs[refs.length - 1].text += ' ' + b.text.trim(); refs[refs.length - 1].blocks.push(b); }
        else refs.push({ idx: b.idx, text: b.text, blocks: [b] });
      }
      stats.refs = refs.length;
      const wiki = B.filter((b) => b.type === 'p' && /wikipedia|ウィキペディア/i.test(b.text));
      wiki.forEach((b) => {
        const m = b.text.match(/wikipedia|ウィキペディア/i);
        add({ level: 'ng', key: 'refs', block: b.idx, spans: [[m.index, m.index + m[0].length]], msg: 'Wikipediaは引用・参考文献に使いません', fix: 'Wikipediaの記事の元になった本・論文・公的機関のサイトを探して、そちらを使おう' });
      });
      let bad = 0;
      let expectNo = 1;
      const anyNumbered = refs.some((b) => NUM.test(b.text));
      const noPeriod = [];
      for (const b of refs) {
        const t = b.text.trim();
        const z = z2h(t);
        const probs = [];
        let lvl = 'warn';
        const numM = z.match(/^[\(（]\s*(\d{1,2})\s*[\)）]/);
        if (!numM) { if (anyNumbered) probs.push(['番号がありません', '行の始めに(1)(2)…と番号を']); expectNo++; }
        else { if (+numM[1] !== expectNo) probs.push([`番号が(${expectNo})ではなく(${numM[1]})になっています`, '番号を順にそろえよう']); expectNo = +numM[1] + 1; }
        const url = /https?:\/\/|www\./i.test(t);
        const foreign = !JA.test(t.replace(/^[\(（]?\s*[0-9０-９]*\s*[\)）]?/, ''));
        if (/file:\/\/|[A-Z]:[\\\/]Users|\/Users\/|Downloads[\\\/]/i.test(t)) { probs.push(['自分のパソコンの中のファイルの場所になっています（ほかの人は開けません）', 'その資料が載っているサイトのURLか、本・雑誌の情報を']); lvl = 'ng'; }
        const news = /新聞/.test(t) && !/『[^』]*新聞[^』]*研究/.test(t);
        if (url) {
          if (JA.test(t) && !/[「『“"]/.test(t)) probs.push(['ページ（記事）の題名に「」が付いていません', '「書いた人・団体名「記事名」（URL 見た日）」の形に']);
          if (!/((19|20)\d{2}\s*[年\/\.\-]\s*\d{1,2}\s*[月\/\.\-]\s*\d{1,2}).{0,8}(取得|閲覧|参照|アクセス|確認)|(取得|閲覧|参照|アクセス|確認)[日:：\s]*(19|20)\d{2}/.test(z)) probs.push(['サイトを見た日が書いてありません', '最後に「2026年9月1日情報取得」のように見た日を']);
          if (/^[\(（]?\s*\d*\s*[\)）]?\s*https?:/.test(z)) { probs.push(['URLだけになっています', '書いた人・団体名と「記事名」を、URLの前に']); lvl = 'ng'; }
        } else if (news) {
          if (!/(19|20)\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/.test(z)) probs.push(['新聞の日付がありません', '「〇年〇月〇日付朝刊」のように']);
        } else {
          if (!/[『「“"]/.test(t) && JA.test(t)) { probs.push(['題名に『』や「」が付いていません', '本の題は『』、論文・記事の題は「」でくくろう']); lvl = 'ng'; }
          if (/「[^」]+」/.test(t) && !/『/.test(t) && !foreign && /(学会|研究|紀要|学報|論集|雑誌|学|誌|力学|経済)/.test(t.slice(t.lastIndexOf('」') + 1))) probs.push(['載っている雑誌・本の名前に『』が付いていません', '論文の題は「」、載っている雑誌・本の名前は『』でくくろう']);
          if (/\d+\s*[-‐－〜~]\s*\d+\s*項|\d+\s*項/.test(z) && !/条\s*\d+\s*項/.test(z)) probs.push(['ページが「項」になっています', '「頁」か「pp.53-72」の形に']);
          const yParen = /^[\(（]?\s*\d*\s*[\)）]?[^『「（(]{1,40}[（(]\s*(1[89]|20)\d{2}[a-z]?\s*[）)]/.test(z);
          if (!yParen && !foreign) {
            if (/(1[89]|20)\d{2}/.test(z)) probs.push(['刊行年の位置が見本と違います', '著者名のすぐ後に（2000）の形で']);
            else { probs.push(['刊行年がありません', '著者名のすぐ後に（2000）の形で']); lvl = 'ng'; }
          }
          if (/「/.test(t) && /『/.test(t) && !/pp?\.\s*\d|\d+\s*[-‐－〜~]\s*\d+\s*[頁項]|\d+\s*[頁項]/.test(z)) probs.push(['載っているページがありません', '「pp.24-25」のように']);
          const lastClose = Math.max(t.lastIndexOf('』'), t.lastIndexOf('」'));
          if (lastClose >= 0) {
            const after = t.slice(lastClose + 1).replace(/pp?\.\s*[\d\-‐－〜~]+|[\d\(\)（）]+|[\s　,，、.．。]/g, '');
            // 大学の紀要・学会誌などの雑誌は出版社がないことがあるので、本（『』が雑誌らしくないもの）だけ出版社を求める
            const media = (t.match(/『([^』]+)』/) || [])[1] || '';
            const journal = /紀要|研究|論集|論文集|講演|大会|学会|大学|白書|論叢|学報|年報|報告|雑誌|学会誌|会誌|学誌|ジャーナル|レビュー|Journal|Bulletin|Review|Proceedings|Annals|Studies|第\s*[0-9０-９]+\s*[巻号]|Vol\.?|No\./i.test(media + ' ' + t.slice(t.indexOf(media)));
            if (after.length < 2 && media && !journal) probs.push(['出版社が書いてありません', '最後に出版社名を（本の場合）']);
          }
        }
        if (!/[.．。]$/.test(t)) noPeriod.push(b);
        if (probs.length) {
          bad++;
          add({ level: probs.length > 1 && lvl !== 'ng' ? 'warn' : lvl, key: 'refs', block: b.idx, msg: probs.map((p) => p[0]).join('。'), fix: probs.map((p) => p[1]).join('／') });
        }
      }
      const rh = byRole('refhead')[0];
      if (refs.length && !anyNumbered) { add({ level: 'warn', key: 'refs', block: rh ? rh.idx : refs[0].idx, msg: `文献に番号がありません（${refs.length}件）`, fix: '見本のように、それぞれの始めに(1)(2)…と番号を付け、本文でもその番号で示そう' }); }
      if (noPeriod.length) add({ level: noPeriod.length === refs.length ? 'warn' : 'hint', key: 'refs', block: rh ? rh.idx : noPeriod[0].idx, msg: `最後に「.」がない文献が${noPeriod.length}件あります`, fix: '見本のように、それぞれの最後を「.」で終えよう' });
      if (!refs.length && byRole('refhead').length) add({ level: 'ng', key: 'refs', block: byRole('refhead')[0].idx, msg: '文献が1つも書かれていません', fix: '使った本・論文・サイトを見本の形で並べよう' });
      const st = !refs.length || wiki.length || issues.some((x) => x.key === 'refs' && x.level === 'ng') ? 'ng' : bad ? 'warn' : 'ok';
      setCheck('refs', st, refs.length ? `${refs.length}件${bad ? `（うち${bad}件に直すところ）` : ''}` : '文献がありません');
    }

    // ---- 16 本文中の出典の示し方 ----
    {
      const bodies = byRole('body');
      const cite = bodies.some((b) => b.footnotes > 0 || /[\(（][0-9０-９]{1,2}[\)）]|\[[0-9０-９]{1,2}\]|注[0-9０-９]|[0-9０-９]{1,2}[)）](?=[、。]|$)|[（(][^（）()]{1,25}[\s,，、]?(19|20)\d{2}[a-z]?[:：]?[^（）()]*[）)]|[一-龥ぁ-んァ-ヶーA-Za-z・]{1,20}(ら|他|ほか|et al\.?)?\s*[（(]\s*(19|20)\d{2}[a-z]?\s*[）)]/.test(b.text));
      if (!byRole('ref').length) setCheck('cite', 'warn', '文献がないため判定できません');
      else if (!cite) {
        setCheck('cite', 'warn', '本文に文献番号が見当たりません');
        const b = byRole('heading').find((h) => /論証/.test(h.text)) || bodies[0];
        add({ level: 'warn', key: 'cite', block: b ? b.idx : null, msg: '本文のどこで文献を使ったかが分かりません', fix: '引用・参考にした文の後に文献の番号(1)などを付けよう（事実と意見、自分の説と他人の説を書き分けるため）' });
      } else setCheck('cite', 'ok', '文献番号あり');
    }

    // ---- 17 図表 ----
    {
      const figs = B.filter((b) => b.images.length && b.idx > d.headerEnd && b.role !== 'ref');
      const tbls = byRole('table').filter((b) => b.idx > d.headerEnd);
      const ne = (i, dir) => { let j = i + dir; while (B[j] && B[j].role === 'empty') j += dir; return B[j] || null; };
      let n = 0;
      stats.figures = figs.length; stats.tables = tbls.length;
      const hasSource = (from, k) => { let x = from; for (let i = 0; i < k && x; i++) { if (/出典|出所|出展|資料[)）:：]|より作成|作成[)）]/.test(x.text)) return true; x = ne(x.idx, 1); } return false; };
      for (const f of figs) {
        const below = ne(f.idx, 1), above = ne(f.idx, -1);
        const inSame = RE_CAPTION.test(f.text) && /図|グラフ|写真|資料|Fig/i.test(f.text);
        const capBelow = inSame || (below && below.role === 'caption' && /図|グラフ|写真|資料|Fig/i.test(below.text));
        const capAbove = above && above.role === 'caption' && /図|グラフ|写真|資料|Fig/i.test(above.text);
        const src = hasSource(capBelow ? (inSame ? f : below) : f, 3);
        if (!capBelow && capAbove) { n++; add({ level: 'ng', key: 'figure', block: above.idx, msg: '図の番号・題が図の上にあります' + (src ? '' : '。出典もありません'), fix: '図の番号と題（図-1 〇〇）は図の下に' + (src ? '' : '。その下に「出典）〇〇」') }); }
        else if (!capBelow) { n++; add({ level: 'warn', key: 'figure', block: f.idx, msg: src ? '図に番号と題がありません' : '図に番号・題と出典がありません', fix: src ? '図の下に「図-1 〇〇」と番号と題を' : '図の下に「図-1 〇〇」と番号と題、その下に「出典）〇〇『〇〇』（年）」' }); }
        else if (!src) { n++; add({ level: 'warn', key: 'figure', block: inSame ? f.idx : below.idx, msg: '図の出典がありません', fix: '図の題の下に「出典）〇〇『〇〇』（年）」のように、元にした資料を' }); }
      }
      for (const t of tbls) {
        const below = ne(t.idx, 1), above = ne(t.idx, -1);
        const capAbove = above && above.role === 'caption' && /表|Table/i.test(above.text);
        const capBelow = below && below.role === 'caption' && /表|Table/i.test(below.text);
        const src = hasSource(below, 2);
        if (!capAbove && capBelow) { n++; add({ level: 'ng', key: 'figure', block: below.idx, msg: '表の番号・題が表の下にあります', fix: '表の番号と題（表-1 〇〇）は表の上に' }); }
        else if (!capAbove) { n++; add({ level: 'warn', key: 'figure', block: t.idx, msg: src ? '表に番号と題がありません' : '表に番号・題と出典がありません', fix: src ? '表の上に「表-1 〇〇」と番号と題を' : '表の上に「表-1 〇〇」と番号と題、表の下に「出典）〇〇」' }); }
        else if (!src) { n++; add({ level: 'warn', key: 'figure', block: t.idx, msg: '表の出典がありません', fix: '表の下に「出典）〇〇」と、元にした資料を' }); }
      }
      // 通し番号
      const capNums = { 図: [], 表: [] };
      // 1行に「図1　図2」と並べて書く場合もあるので、行の中の番号をすべて拾う
      byRole('caption').forEach((c) => { for (const m of z2h(c.text).matchAll(/(図|表)\s*[-－‐–—ー]?\s*(\d+)(?:[-－‐](\d+))?/g)) if (!m[3]) capNums[m[1]].push({ n: +m[2], b: c }); });
      // 「出展」は「出典」の書き間違い
      B.filter((b) => b.type === 'p' && /出展/.test(b.text) && !/展示|展覧/.test(b.text)).forEach((b) => { const i = b.text.indexOf('出展'); n++; add({ level: 'warn', key: 'figure', block: b.idx, spans: [[i, i + 2]], msg: '「出展」になっています', fix: '「出典」（資料の出どころ）に直そう' }); });
      for (const k of ['図', '表']) capNums[k].forEach((x, i) => { if (x.n !== i + 1) { n++; add({ level: 'warn', key: 'figure', block: x.b.idx, msg: `${k}の番号が通し番号になっていません（${i + 1}番目なのに${x.n}）`, fix: `${k}-1、${k}-2…と順に` }); } });
      const ngF = issues.some((x) => x.key === 'figure' && x.level === 'ng');
      setCheck('figure', !figs.length && !tbls.length ? 'none' : ngF ? 'ng' : n ? 'warn' : 'ok', !figs.length && !tbls.length ? '図表なし' : `図${figs.length}・表${tbls.length}${n ? `／${n}か所` : ''}`);
    }

    // ---- 18 文字の色・蛍光ペン ----
    {
      let n = 0;
      for (const b of B) {
        if (b.type !== 'p' || b.idx <= -1) continue;
        const colored = (r) => (r.r.color && !/^(auto|000000)$/i.test(r.r.color)) || (r.r.highlight && r.r.highlight !== 'none') || r.r.shade;
        const links = b.runs.filter((r) => norm(r.text) && r.r.link && colored(r));
        if (links.length) {
          n++;
          add({ level: 'warn', key: 'color', block: b.idx, spans: links.map((r) => [r.start, r.end]), msg: 'URLが青い下線のリンクのままです', fix: 'URLを右クリック→［ハイパーリンクの削除］で黒い文字に' });
        }
        const bad = b.runs.filter((r) => norm(r.text) && !r.r.link && !inLeftover(b, r) && colored(r));
        if (bad.length) {
          n++;
          add({ level: 'warn', key: 'color', block: b.idx, spans: bad.map((r) => [r.start, r.end]), msg: '文字に色・蛍光ペン・網かけが付いています', fix: '文字はすべて黒にしよう（ネットから貼り付けたときに残りやすい）' });
        }
      }
      setCheck('color', n ? 'warn' : 'ok', n ? `${n}段落` : 'すべて黒');
    }

    // 言葉づかいの助言の数
    const talkN = issues.filter((x) => x.key === 'talk').length;
    setCheck('talk', talkN ? 'hint' : 'ok', talkN ? `${talkN}か所` : '話し言葉は見当たりません');

    stats.bodyChars = B.filter((b) => ['body', 'heading'].includes(b.role)).reduce((s, b) => s + countChars(b.text), 0);

    // 読む順に並べ、番号を振る
    const order = (x) => (x.block === null || x.block === undefined ? -1 : x.block) * 1000 + (x.spans && x.spans.length ? Math.min(...x.spans.map((s) => s[0])) / 1000 : 0);
    issues.sort((a, b) => order(a) - order(b));
    let no = 0;
    issues.forEach((x) => { x.no = ++no; });
    return { doc: d, issues, checks, stats };
  }

  // ---------------- ページ数の目安 ----------------
  // 半角（英数字・URL）は全角の半分の幅
  const textWidth = (s) => [...s].reduce((w, c) => w + (/[\x20-\x7e\uff61-\uff9f]/.test(c) ? 0.5 : 1), 0);

  function estimatePages(d) {
    const B = d.blocks;
    const bodySec = d.sections.find((s) => s.cols >= 2) || d.sections[d.sections.length - 1];
    const headSec = d.sections[0];
    const pageH = (bodySec.h || RULE.pageH);
    const pageW = (bodySec.w || RULE.pageW);
    const mar = bodySec.mar || { t: 851, b: 851, l: 851, r: 851 };
    const usableH = pageH - mar.t - mar.b;
    const cpl = (sz, w) => Math.max(10, Math.floor(w / (sz * 10)));
    // 見出し部分の高さ
    const hm = headSec.mar || mar;
    const fullW = (headSec.w || pageW) - hm.l - hm.r;
    let headTw = 0;
    for (const b of B) {
      if (b.idx > d.headerEnd) break;
      const sz = (b.runs[0] && b.runs[0].r.sz) || 18;
      const lines = Math.max(1, Math.ceil(textWidth(b.text) / cpl(sz, fullW)));
      headTw += lines * Math.max(headSec.linePitch || 360, sz * 10 * 1.5);
    }
    // セクションごとの紙の形（横向きの表のページなど）で1ページに入る行数を出し、ページ数を足していく
    const geom = (sec) => {
      const S = sec || bodySec;
      const m = S.mar || mar;
      const c = Math.max(1, S.cols || 1);
      const w = ((S.w || pageW) - m.l - m.r - (c - 1) * (S.colSpace || 425)) / c;
      const p = Math.max(S.linePitch || 360, 1);
      return { cols: c, colW: w, pitch: p, perPage: Math.max(1, Math.floor(((S.h || pageH) - m.t - m.b) / p)) * c };
    };
    let pages = headTw / usableH;
    let prevSec = null;
    for (const b of B) {
      if (b.idx <= d.headerEnd) continue;
      const sec = d.sections[b.sec];
      if (prevSec !== null && b.sec !== prevSec && sec && sec.type !== 'continuous') pages = Math.ceil(pages - 0.001);
      prevSec = b.sec;
      const g = geom(sec);
      let l = 0;
      if (b.type === 'table') {
        // 表の幅（列の設定）と、行の高さの指定を使う
        const tw = b.gridW > 0 ? Math.min(b.gridW, g.colW * g.cols + 480 * (g.cols - 1)) : g.colW;
        const perLine = Math.max(3, cpl(b.sz || 18, tw));
        let tw2 = 0;
        b.rows.forEach((r, i) => {
          const cw = perLine / Math.max(1, r.length);
          const ln = Math.max(1, ...r.map((c) => c.split('\n').reduce((x, seg) => x + Math.max(1, Math.ceil(textWidth(seg) / Math.max(3, cw))), 0)));
          tw2 += Math.max(b.heights[i] || 0, ln * (b.sz || 18) * 10 * 1.4 + 100);
        });
        l = tw2 / g.pitch + 1;
        // 段組みの中でも表は段の幅を超えて置けないので、幅の広い表は1段分として数える
        if (b.gridW > g.colW * 1.2 && g.cols > 1) l *= g.cols;
      } else {
        const sz = (b.runs.find((r) => norm(r.text)) || { r: { sz: 18 } }).r.sz || 18;
        for (const seg of b.text.split('\n')) l += Math.max(1, Math.ceil((textWidth(seg) + ((b.ind.firstLineChars || 0) / 100)) / cpl(sz, g.colW)));
        if (sz > 18) l = Math.ceil(l * Math.max(1, (sz * 10 * 1.4) / g.pitch));
        for (const im of b.images) {
          // 浮かせた（文字列の折り返しあり）画像は、まわりの空行で場所を空けていることが多いので高さに数えない
          if (!im.cy || im.anchor) continue;
          const scale = im.cx > 0 ? Math.min(1, (g.colW / 1440 * 914400) / im.cx) : 1;
          l += Math.ceil((im.cy * scale) / 635 / g.pitch);
        }
      }
      if (b.pageBreakBefore) pages = Math.ceil(pages - 0.001);
      pages += l / g.perPage;
      if (b.pageBreak) pages = Math.ceil(pages - 0.001);
    }
    // 学校のWindows版Wordで保存された卒論12本（どれも4ページ）に合わせた補正
    const K = window.AKAPEN_K || 0.86;
    return Math.max(0.3, pages * K);
  }

  global.Akapen = { readDocx, analyze, RULE, countChars };
})(window);
