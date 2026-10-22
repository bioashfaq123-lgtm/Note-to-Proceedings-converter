/* Reads a Note (.docx) and pulls out the pieces a Proceeding needs. */
(function (NP) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  // Known acronyms -> full programme names (edit freely; used when a note mentions them)
  NP.ACRONYMS = {
    PAMP: 'Participatory Agricultural Management Programme',
    RAWE: 'Rural Agricultural Work Experience',
    ELP: 'Experiential Learning Programme'
  };

  const clean = s => s.replace(/[\u00a0\t ]+/g, ' ').replace(/ *\n */g, '\n').trim();

  function walkText(node, out) {
    for (const c of node.childNodes) {
      if (c.nodeType !== 1 || c.namespaceURI !== W) continue;
      const n = c.localName;
      if (n === 't') out.push(c.textContent);
      else if (n === 'br' || n === 'cr') out.push('\n');
      else if (n === 'tab') out.push(' ');
      else if (n === 'del') continue;                 // deleted tracked text
      else walkText(c, out);
    }
  }
  function paraText(p) { const o = []; walkText(p, o); return clean(o.join('')); }
  function paraBold(p) {
    const runs = [...p.getElementsByTagNameNS(W, 'r')].filter(r => r.getElementsByTagNameNS(W, 't').length && r.textContent.trim());
    return runs.length > 0 && runs.every(r => {
      const rp = r.getElementsByTagNameNS(W, 'rPr')[0];
      return rp && rp.getElementsByTagNameNS(W, 'b').length;
    });
  }
  function cellLines(tc) {
    const lines = [];
    for (const p of tc.getElementsByTagNameNS(W, 'p')) {
      paraText(p).split('\n').forEach(l => { if (l.trim()) lines.push(l.trim()); });
    }
    return lines;
  }

  NP.normalizeDesignation = function (s) {
    return s.replace(/Asst\.?\s*Prof\.?/i, 'Asst. Prof.')
      .replace(/Assoc\.?\s*Prof\.?/i, 'Assoc. Prof.')
      .replace(/\s*\(\s*Hort\.?\s*\)/i, ' (Horti.)')
      .replace(/Prof\.(?=\()/, 'Prof. ')
      .replace(/\s+/g, ' ').trim();
  };

  const capFirst = s => s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  const lowFirst = s => s ? s.charAt(0).toLowerCase() + s.slice(1) : s;
  NP.capFirst = capFirst; NP.lowFirst = lowFirst;

  function detectPurpose(body) {
    let core = '', place = '', acro = '';
    let m = body.match(/purchase of (?:the )?(?:following )?(.+?)(?=\s+(?:from|at|for|to be|which|as follows)\b|[.;:]|$)/i);
    if (m) core = 'purchase of ' + m[1].trim();
    else {
      m = body.match(/preparation of (?:various )?(.+?)(?=\s+(?:like|such as)\b|[.,;]|$)/i);
      if (m) core = 'purchase of ' + m[1].trim();
    }
    const pm = body.match(/\b(?:in|for|at)\s+(?:(?:the|existing)\s+)*([A-Za-z]+(?:\s[A-Za-z]+)?\s(?:orchard|garden|field|farm|nursery|laboratory|lab|hostel|campus|unit))\b/i);
    if (pm) place = pm[1].trim();
    for (const code of Object.keys(NP.ACRONYMS)) {
      if (new RegExp('\\b' + code + '\\b').test(body)) { acro = NP.ACRONYMS[code] + ' (' + code + ')'; break; }
    }
    if (!core) return '';
    let out = capFirst(core);
    if (place) out += ' for ' + place;
    if (acro) out += ' for ' + acro;
    return out;
  }

  function detectHeadLines(lines) {
    // lines from the note after "head of account"; normalise quotes / dashes
    const L = lines.map(l => l.replace(/[“”"]/g, '').replace(/\s*[-–—]\s*/g, ' – ').replace(/\s+/g, ' ').trim()).filter(Boolean);
    let code = '', fund = '';
    for (const l of L) {
      if (/^\d{5,7}\b/.test(l)) code = code || l;
      else fund = fund || l;
    }
    if (/revolving fund/i.test(fund)) fund = 'Revolving Fund';
    else fund = fund.replace(/\s*–\s*APT\.?,?\s*Jagtial\.?/i, ', APT., Jagtial').replace(/^(\w+) – /, '$1, ');
    return { code, fund };
  }

  NP.parseNote = async function (arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const f = zip.file('word/document.xml');
    if (!f) throw new Error('This does not look like a .docx file.');
    const xml = await f.async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const body = doc.getElementsByTagNameNS(W, 'body')[0];

    const blocks = [];
    for (const el of body.children) {
      if (el.localName === 'p') blocks.push({ type: 'p', text: paraText(el), bold: paraBold(el) });
      else if (el.localName === 'tbl') {
        const rows = [...el.getElementsByTagNameNS(W, 'tr')].map(tr =>
          [...tr.childNodes].filter(c => c.localName === 'tc').map(tc => ({ lines: cellLines(tc) })));
        blocks.push({ type: 'tbl', rows });
      }
    }
    const warnings = [];
    const textOf = c => c.lines.join(' ');

    // ---- items table
    let itemsIdx = blocks.findIndex(b => b.type === 'tbl' && b.rows.length > 1 && b.rows[0].length >= 4 &&
      /particular|product|item|description|amount/i.test(b.rows[0].map(textOf).join(' ')));
    const items = []; let statedTotal = NaN, tableHeading = 'Particulars';
    if (itemsIdx < 0) warnings.push('Could not find the items table in the note — add the items manually below.');
    else {
      const rows = blocks[itemsIdx].rows;
      const head = rows[0].map(textOf);
      const col = (re, d) => { const i = head.findIndex(h => re.test(h)); return i < 0 ? d : i; };
      const iName = col(/particular|product|item|description/i, 1), iQty = col(/qty|quant/i, 2),
        iRate = col(/rate/i, 3), iAmt = col(/amount/i, 4);
      if (/product/i.test(head[iName] || '')) tableHeading = 'Product Name';
      for (const r of rows.slice(1)) {
        const t = r.map(textOf);
        if (/^total/i.test(t[0] || '') || /^total/i.test(t[iName] || '')) {
          const nums = t.map(NP.parseNum).filter(n => !isNaN(n));
          if (nums.length) statedTotal = nums[nums.length - 1];
          continue;
        }
        if (!t.some(x => x.trim())) continue;
        items.push({ name: t[iName] || '', qty: t[iQty] || '', rate: t[iRate] || '', amount: t[iAmt] || '' });
      }
    }

    // ---- paragraphs before / after the table
    const before = blocks.slice(0, itemsIdx < 0 ? blocks.length : itemsIdx).filter(b => b.type === 'p' && b.text);
    const after = itemsIdx < 0 ? [] : blocks.slice(itemsIdx + 1);

    // date
    let noteDate = '';
    for (const b of before.slice(0, 4)) {
      const m = b.text.match(/D(?:t|ate)\s*[:.]?\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{4})/i);
      if (m) { noteDate = m[1].padStart(2, '0') + '.' + m[2].padStart(2, '0') + '.' + m[3]; break; }
    }
    if (!noteDate) warnings.push('Note date (Dt: …) was not found — please enter it.');

    // main paragraph(s)
    let sirIdx = before.findIndex(b => /^sir\b/i.test(b.text));
    const mainParas = before.slice(sirIdx >= 0 ? sirIdx + 1 : 1).map(b => b.text);
    const noteBody = mainParas.join(' ');

    // head of account
    const hIdx = after.findIndex(b => b.type === 'p' && /head of account/i.test(b.text));
    const headRaw = hIdx < 0 ? [] : after.slice(hIdx + 1).filter(b => b.type === 'p' && b.text).map(b => b.text);
    const head = detectHeadLines(headRaw);
    if (!head.fund && !head.code) warnings.push('Head of account was not found — please fill it in.');

    // signatory: last table after the items table, else last paragraphs
    let sigLines = [];
    const sigTbl = [...after].reverse().find(b => b.type === 'tbl');
    if (sigTbl) sigLines = sigTbl.rows.flat().flatMap(c => c.lines);
    sigLines = sigLines.map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const authorName = sigLines[0] || '';
    const authorDesignation = NP.normalizeDesignation(sigLines[1] || '');
    if (!authorDesignation) warnings.push('Could not read the signatory designation — please fill it in.');

    // amount sanity check
    const sum = items.reduce((a, it) => a + (isNaN(NP.parseNum(it.amount)) ? 0 : NP.parseNum(it.amount)), 0);
    if (!isNaN(statedTotal) && items.length && Math.abs(sum - statedTotal) > 0.005)
      warnings.push(`The note's stated total (Rs. ${statedTotal}) differs from the sum of the item amounts (Rs. ${sum}). Please check the figures.`);
    if (isNaN(statedTotal)) {
      const m = noteBody.match(/Rs\.?\s*([\d,]+(?:\.\d+)?)/i); if (m) statedTotal = NP.parseNum(m[1]);
    }

    const purpose = detectPurpose(noteBody);
    if (!purpose) warnings.push('Could not work out the purpose for the subject line — please type it in the “Purpose” box.');

    return { noteDate, noteBody, items, tableHeading, statedTotal, authorName, authorDesignation,
      headCode: head.code, headFund: head.fund, purpose, warnings };
  };
})(window.NP = window.NP || {});
