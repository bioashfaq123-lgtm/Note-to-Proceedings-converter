/* UI logic: upload -> parse -> review form + live preview -> download .docx */
(function () {
  const NP = window.NP;
  const $ = s => document.querySelector(s);
  const FACTORY = {
    university: 'PROF. JAYASHANKAR TELANGANA AGRICULTURAL UNIVERSITY',
    officeTitle: 'Proceedings of the Principal,',
    officeAddr: 'Agricultural Polytechnic, Jagtial - 505 529',
    officeShort: 'Agricultural Polytechnic, Jagtial',
    present: 'Dr. A. Venkat Reddy',
    procSuffix: '/A/APT-JGL/2026-27',
    delegation: '2(viii)',
    headCode: '121541 – Agricultural Polytechnic, Jagtial',
    refs: ['University Proc. No. 551/S.C./2011.  Dated: 15-07-2011',
      'PJTSAU Cir. Memo. No. 13388/SC/A1/2011, Dated: 08-08-2011'],
    ccLines: ['CC to Bills Section of APT., Polasa, Jagtial.', 'Copy to Deputy Comptroller, RAO, RARS., Jagtial.']
  };

  // ---------- storage (safe if blocked)
  const store = {
    get(k, d) { try { const v = localStorage.getItem('np.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('np.' + k, JSON.stringify(v)); } catch (e) { /* ignore */ } }
  };
  const settings = () => Object.assign({}, FACTORY, store.get('settings', {}));

  let S = null, fileName = '';
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const toast = (m, ms = 3200) => { const t = $('#toast'); t.textContent = m; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => t.hidden = true, ms); };
  const todayISO = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  const isoToDMY = iso => iso ? iso.split('-').reverse().join('-') : '';
  const dmyToISO = dmy => { const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(dmy || ''); return m ? `${m[3]}-${m[2]}-${m[1]}` : ''; };

  $('#logo').src = 'data:image/jpeg;base64,' + window.LOGO_BASE64;

  // ---------- upload
  const drop = $('#drop'), fileIn = $('#file');
  fileIn.addEventListener('change', () => fileIn.files[0] && handleFile(fileIn.files[0]));
  drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileIn.click(); } });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => e.dataTransfer.files[0] && handleFile(e.dataTransfer.files[0]));

  async function handleFile(file) {
    const err = $('#uploadError'); err.hidden = true;
    if (!/\.docx$/i.test(file.name)) { err.textContent = 'Please choose a Word file with the .docx extension (open .doc files in Word and “Save As” .docx).'; err.hidden = false; return; }
    try {
      const parsed = await NP.parseNote(await file.arrayBuffer());
      fileName = file.name;
      startReview(parsed);
    } catch (e) {
      console.error(e);
      err.textContent = 'Sorry, this file could not be read: ' + e.message; err.hidden = false;
    } finally { fileIn.value = ''; }
  }

  // ---------- build state from parsed note + saved defaults
  function startReview(p) {
    const cfg = settings();
    const last = parseInt(store.get('lastProcNo', ''), 10);
    S = {
      university: cfg.university, officeTitle: cfg.officeTitle, officeAddr: cfg.officeAddr, officeShort: cfg.officeShort,
      present: cfg.present, procSuffix: cfg.procSuffix, delegation: cfg.delegation,
      procNo: isNaN(last) ? '' : String(last + 1), procDate: isoToDMY(todayISO()),
      noteDate: p.noteDate, authorDesignation: p.authorDesignation, authorName: p.authorName,
      items: p.items.length ? p.items : [{ name: '', qty: '', rate: '', amount: '' }],
      tableHeading: p.tableHeading, purpose: p.purpose, vendor: '',
      sanctionTo: p.authorDesignation ? `the ${p.authorDesignation}, Agricultural Polytechnic, Jagtial` : '',
      headCode: p.headCode || cfg.headCode, headFund: p.headFund, headExtra: '',
      includePayment: true, refs: cfg.refs.slice(), ccLines: cfg.ccLines.slice(),
      toLines: p.authorDesignation ? [`The ${p.authorDesignation}, APT., Jagtial.`] : [],
      subjectManual: null, bodyManual: null
    };
    $('#fileName').textContent = 'From: ' + fileName;
    $('#viewUpload').hidden = true; $('#viewReview').hidden = false;
    const w = p.warnings.slice();
    if (!S.procNo) w.push('Enter the Proc. No. (after your first download the next number is suggested automatically).');
    $('#warnings').innerHTML = w.length ? `<div class="alert warn"><b>Please check:</b><ul>${w.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : '';
    fillForm(); renderItems(); renderRefs(); refresh();
    window.scrollTo(0, 0);
  }

  // ---------- form <-> state
  const textAreaKeys = { toLinesText: 'toLines', ccLinesText: 'ccLines' };
  function fillForm() {
    document.querySelectorAll('#form [data-k]').forEach(el => {
      const k = el.dataset.k;
      if (textAreaKeys[k]) el.value = S[textAreaKeys[k]].join('\n');
      else if (el.type === 'checkbox') el.checked = !!S[k];
      else el.value = S[k] == null ? '' : S[k];
    });
    $('#procDateInput').value = dmyToISO(S.procDate);
    $('#vendorList').innerHTML = store.get('vendors', []).map(v => `<option value="${esc(v)}">`).join('');
  }

  $('#form').addEventListener('input', e => {
    const el = e.target, k = el.dataset && el.dataset.k;
    if (el.id === 'procDateInput') { S.procDate = isoToDMY(el.value); }
    else if (el.id === 'subjectBox') { S.subjectManual = el.value; }
    else if (el.id === 'bodyBox') { S.bodyManual = el.value; }
    else if (k) {
      if (textAreaKeys[k]) S[textAreaKeys[k]] = el.value.split('\n');
      else S[k] = el.type === 'checkbox' ? el.checked : el.value;
    }
    refresh(el);
  });
  document.querySelectorAll('[data-reset]').forEach(b => b.addEventListener('click', () => {
    if (b.dataset.reset === 'subject') S.subjectManual = null; else S.bodyManual = null;
    refresh();
  }));

  // ---------- items
  function renderItems() {
    const tb = $('#itemsTable tbody');
    tb.innerHTML = S.items.map((it, i) => `<tr>
      <td class="n">${i + 1}</td>
      <td><input data-i="${i}" data-f="name" value="${esc(it.name)}"></td>
      <td><input data-i="${i}" data-f="qty" value="${esc(it.qty)}" size="6"></td>
      <td><input data-i="${i}" data-f="rate" value="${esc(it.rate)}" size="6"></td>
      <td><input data-i="${i}" data-f="amount" value="${esc(it.amount)}" size="8"></td>
      <td><button class="x" type="button" data-del="${i}" title="Remove row">✕</button></td></tr>`).join('');
  }
  $('#itemsTable').addEventListener('input', e => {
    const el = e.target; if (el.dataset.i == null) return;
    S.items[+el.dataset.i][el.dataset.f] = el.value; refresh(el);
  });
  $('#itemsTable').addEventListener('click', e => {
    if (e.target.dataset.del == null) return;
    S.items.splice(+e.target.dataset.del, 1); if (!S.items.length) S.items.push({ name: '', qty: '', rate: '', amount: '' });
    renderItems(); refresh();
  });
  $('#btnAddItem').addEventListener('click', () => { S.items.push({ name: '', qty: '', rate: '', amount: '' }); renderItems(); refresh(); });

  // ---------- references
  function renderRefs() {
    $('#refsList').innerHTML = S.refs.map((r, i) => `<div class="ref-row"><span>${i + 1}.</span>
      <input data-r="${i}" value="${esc(r)}"><button class="x" type="button" data-rdel="${i}" title="Remove">✕</button></div>`).join('');
  }
  $('#refsList').addEventListener('input', e => { if (e.target.dataset.r != null) { S.refs[+e.target.dataset.r] = e.target.value; refresh(e.target); } });
  $('#refsList').addEventListener('click', e => { if (e.target.dataset.rdel != null) { S.refs.splice(+e.target.dataset.rdel, 1); renderRefs(); refresh(); } });
  $('#btnAddRef').addEventListener('click', () => { S.refs.push(''); renderRefs(); refresh(); });

  // ---------- refresh derived UI
  function refresh(active) {
    const V = NP.derive(S);
    $('#totalCell').textContent = 'Rs. ' + V.total.toFixed(2);
    if (active !== $('#subjectBox')) $('#subjectBox').value = V.subject;
    if (active !== $('#bodyBox')) $('#bodyBox').value = V.body;
    ['procNo', 'purpose', 'vendor'].forEach(k => { const el = document.querySelector(`[data-k="${k}"]`); if (el) el.classList.toggle('need', !String(S[k] || '').trim()); });
    $('#preview').innerHTML = previewHTML(S, V);
  }

  const md = s => esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/(\d+)(st|nd|rd|th)\b/g, '$1<sup>$2</sup>').replace(/…/g, '<mark>…</mark>');
  function previewHTML(S, V) {
    const rows = S.items.map((it, i) => `<tr><td class="ctr">${i + 1}.</td><td>${esc(it.name)}</td><td class="ctr">${esc(it.qty)}</td><td class="ctr">${esc(it.rate)}</td><td class="ctr">${esc(it.amount)}</td></tr>`).join('');
    const refs = V.allRefs.map((r, i) => `<tr><td>${i === 0 ? 'Ref:' : ''}</td><td>${i + 1}.</td><td>${md(r)}</td></tr>`).join('');
    return `<div class="paper">
      <img class="plogo" src="data:image/jpeg;base64,${window.LOGO_BASE64}" alt="">
      <div class="c"><b>${esc(S.university)}</b></div>
      <div class="c"><b>${esc(S.officeTitle)}</b></div>
      <div class="c"><b>${esc(S.officeAddr)}</b></div>
      <div class="c">&nbsp;</div>
      <div class="c present"><b>Present: ${esc(S.present)}</b></div>
      <div class="procline"><span>${esc(V.procLine)}</span><span>${esc(V.dateLine)}</span></div>
      <table class="sub"><tr><td style="width:46px">Sub:</td><td colspan="2" class="just">${md(V.subject)}</td></tr>${refs}</table>
      <div class="c">***</div>
      ${V.body.split('\n').filter(x => x.trim()).map(t => `<p class="body">${md(t)}</p>`).join('')}
      <table class="items"><tr><th class="ctr">S.No.</th><th class="ctr">${esc(S.tableHeading)}</th><th class="ctr">Qty.(no.)</th><th class="ctr">Rate (Rs.)</th><th class="ctr">Amount (Rs.)</th></tr>
        ${rows}<tr><td colspan="4"><b>Total</b></td><td class="ctr"><b>${V.total.toFixed(2)}</b></td></tr></table>
      <div class="words">${md(V.wordsLine)}</div>
      ${S.includePayment && V.paymentLine ? `<p class="body">${md(V.paymentLine)}</p>` : ''}
      <p class="body">The expenditure shall be debited from the following Head of Account.</p>
      <div class="head">${V.headQuoted.map(esc).join('<br>')}</div>
      <div class="sig">PRINCIPAL<br>${esc(S.officeShort)}</div>
      <div class="tocc">To<br>${S.toLines.concat(S.ccLines).filter(x => x.trim()).map(esc).join('<br>')}</div>
    </div>`;
  }

  // ---------- download
  $('#btnDownload').addEventListener('click', async () => {
    const btn = $('#btnDownload');
    const missing = [];
    if (!String(S.procNo).trim()) missing.push('Proc. No.');
    if (!S.purpose.trim() && S.subjectManual == null) missing.push('Purpose');
    if (missing.length && !confirm('These are still empty: ' + missing.join(', ') + '.\nDownload anyway?')) return;
    btn.disabled = true;
    try {
      const blob = await NP.buildDocx(S);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `Proceeding_${String(S.procNo || 'draft').replace(/[^\w-]+/g, '_')}_${(S.noteDate || 'note').replace(/\./g, '-')}.docx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      const n = parseInt(S.procNo, 10); if (!isNaN(n)) store.set('lastProcNo', n);
      if (S.vendor.trim()) { const v = store.get('vendors', []).filter(x => x !== S.vendor.trim()); v.unshift(S.vendor.trim()); store.set('vendors', v.slice(0, 25)); }
      toast('Proceeding downloaded. Next Proc. No. will be suggested as ' + (isNaN(n) ? '—' : n + 1));
    } catch (e) { console.error(e); alert('Could not create the Word file: ' + e.message); }
    btn.disabled = false;
  });

  $('#btnAnother').addEventListener('click', () => { $('#viewReview').hidden = true; $('#viewUpload').hidden = false; S = null; window.scrollTo(0, 0); });

  $('#btnSaveDefaults').addEventListener('click', () => {
    const cur = store.get('settings', {});
    store.set('settings', Object.assign(cur, {
      refs: S.refs.filter(r => r.trim()), ccLines: S.ccLines.filter(r => r.trim()),
      delegation: S.delegation, present: S.present, procSuffix: S.procSuffix
    }));
    toast('Saved. These will be pre-filled for the next Note.');
  });

  // ---------- office settings dialog
  const dlg = $('#dlgSettings');
  $('#btnSettings').addEventListener('click', () => {
    const cfg = settings();
    dlg.querySelectorAll('[data-s]').forEach(el => el.value = cfg[el.dataset.s]);
    dlg.showModal();
  });
  $('#btnSaveSettings').addEventListener('click', () => {
    const cur = store.get('settings', {});
    dlg.querySelectorAll('[data-s]').forEach(el => cur[el.dataset.s] = el.value);
    store.set('settings', cur); toast('Office settings saved (apply from the next Note).');
  });
  $('#btnResetSettings').addEventListener('click', () => { store.set('settings', {}); dlg.close(); toast('Factory settings restored.'); });
})();
