/* Builds the Proceeding .docx (same layout as the office format) from the reviewed data. */
(function (NP) {
  const FONT = 'Times New Roman', SZ = 24;

  /** "**bold**" markup + auto superscript ordinals (1st, 2nd …) -> TextRun[] */
  NP.makeRuns = function (text, base) {
    const d = window.docx; base = base || {};
    const runs = [];
    String(text || '').split('**').forEach((seg, i) => {
      const bold = base.bold || i % 2 === 1;
      seg.split(/(\d+(?:st|nd|rd|th)\b)/).forEach(part => {
        if (!part) return;
        const m = part.match(/^(\d+)(st|nd|rd|th)$/);
        if (m) {
          runs.push(new d.TextRun({ text: m[1], bold, underline: base.underline, size: SZ, font: FONT }));
          runs.push(new d.TextRun({ text: m[2], bold, superScript: true, size: SZ, font: FONT }));
        } else runs.push(new d.TextRun({ text: part, bold, underline: base.underline, size: SZ, font: FONT }));
      });
    });
    return runs;
  };

  const money = n => NP.inr(n);

  /** Computes every derived value in one place (used by the preview AND the docx). */
  NP.derive = function (S) {
    const total = S.items.reduce((a, it) => a + (isNaN(NP.parseNum(it.amount)) ? 0 : NP.parseNum(it.amount)), 0);
    const amt = money(total);
    const words = NP.amountToWords(total);
    const refs = S.refs.filter(r => r.trim());
    const noteRef = `Note submitted by ${S.authorDesignation || '…'}, APT., Jagtial, Dated: ${S.noteDate || '…'}.`;
    const allRefs = refs.concat([noteRef]);
    const n = refs.length;
    const refPhrase = n === 0 ? '' : n === 1 ? 'reference 1st cited above'
      : n === 2 ? 'references 1st and 2nd cited above' : `references 1st to ${NP.ordinal(n)} cited above`;

    const subjectAuto = `Agricultural Polytechnic, Jagtial – ${S.purpose || '…'} – Expenditure not Exceeding **Rs. ${amt}/-** – Sanction – Orders – Issued.`;
    const bodyPurpose = NP.lowFirst(S.purpose || '…');
    let bodyAuto = (n ? `As per the delegation of powers regulations, 2011 vested under items ${S.delegation}, vide ${refPhrase}, sanction`
      : 'Sanction') + ` is hereby accorded to ${S.sanctionTo || '…'} to incur an expenditure of **Rs. ${amt}/- (${words})** towards ${bodyPurpose}`
      + (S.vendor ? ` from **${S.vendor}**` : '') + ' as detailed below:';
    const wordsLine = `**Rs. ${amt}/- (${words})**`;
    const paymentLine = S.vendor ? `The payment of **Rs. ${amt}/-** shall be made to ${S.vendor}.` : '';
    const head = [S.headCode, S.headFund].map(x => (x || '').trim()).filter(Boolean)
      .concat((S.headExtra || '').split('\n').map(x => x.trim()).filter(Boolean));
    const headQuoted = head.map((l, i) => (i === 0 ? '“' : '') + l + (i === head.length - 1 ? '”' : ''));
    return { total, amt, words, allRefs, noteRef,
      subject: S.subjectManual != null ? S.subjectManual : subjectAuto, subjectAuto,
      body: S.bodyManual != null ? S.bodyManual : bodyAuto, bodyAuto,
      wordsLine, paymentLine, headQuoted,
      procLine: `Proc. No ${S.procNo}${S.procSuffix}`, dateLine: `Dated: ${S.procDate}` };
  };

  NP.buildDocx = async function (S) {
    const d = window.docx, V = NP.derive(S);
    const NONE = { style: d.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
    const LINE = { style: d.BorderStyle.SINGLE, size: 4, color: '000000' };
    const allBorders = { top: LINE, bottom: LINE, left: LINE, right: LINE };
    const P = (children, o) => new d.Paragraph(Object.assign({ children }, o || {}));
    const T = (text, o) => new d.TextRun(Object.assign({ text, size: SZ, font: FONT }, o || {}));
    const blank = () => P([T('')]);

    // logo
    const bin = atob(window.LOGO_BASE64), bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const logo = new d.ImageRun({
      type: 'jpg', data: bytes, transformation: { width: 80, height: 78 },
      floating: {
        horizontalPosition: { relative: d.HorizontalPositionRelativeFrom.COLUMN, offset: 19050 },
        verticalPosition: { relative: d.VerticalPositionRelativeFrom.PARAGRAPH, offset: 3810 },
        wrap: { type: d.TextWrappingType.NONE }, allowOverlap: true
      }
    });

    const kids = [];
    kids.push(P([logo, T(S.university, { bold: true })], { alignment: d.AlignmentType.CENTER, indent: { left: 700 } }));
    kids.push(P([T(S.officeTitle, { bold: true })], { alignment: d.AlignmentType.CENTER }));
    kids.push(P([T(S.officeAddr, { bold: true })], { alignment: d.AlignmentType.CENTER }));
    kids.push(P([T('')], { alignment: d.AlignmentType.CENTER }));
    kids.push(P([T('Present: ' + S.present, { bold: true })], {
      alignment: d.AlignmentType.CENTER,
      border: { bottom: { style: d.BorderStyle.THIN_THICK_THIN_MEDIUM_GAP, size: 18, color: 'auto', space: 0 } }
    }));
    kids.push(P([T(V.procLine, { bold: true, underline: {} }), T('\t'), T(V.dateLine, { bold: true, underline: {} })], {
      tabStops: [{ type: d.TabStopType.RIGHT, position: 9972 }], spacing: { before: 60, after: 60 }
    }));

    // ---- Sub / Ref block (borderless table)
    const cw = [849, 520, 7696];
    const cell = (children, w, o) => new d.TableCell(Object.assign({
      children, width: { size: w, type: d.WidthType.DXA }, borders: noBorders,
      margins: { top: 20, bottom: 20, left: 60, right: 60 }
    }, o || {}));
    const subRows = [new d.TableRow({
      children: [
        cell([P([T('Sub:')])], cw[0]),
        cell(NP.makeRuns(V.subject).length ? [P(NP.makeRuns(V.subject), { alignment: d.AlignmentType.JUSTIFIED, spacing: { line: 276, after: 160 } })] : [blank()],
          cw[1] + cw[2], { columnSpan: 2 })
      ]
    })];
    V.allRefs.forEach((r, i) => subRows.push(new d.TableRow({
      children: [
        cell([P([T(i === 0 ? 'Ref:' : '')])], cw[0]),
        cell([P([T((i + 1) + '.')])], cw[1]),
        cell([P(NP.makeRuns(r), { alignment: d.AlignmentType.LEFT })], cw[2])
      ]
    })));
    kids.push(new d.Table({
      width: { size: 9065, type: d.WidthType.DXA }, columnWidths: cw, alignment: d.AlignmentType.RIGHT,
      borders: Object.assign({ insideHorizontal: NONE, insideVertical: NONE }, noBorders), rows: subRows
    }));
    kids.push(P([T('***')], { alignment: d.AlignmentType.CENTER, spacing: { before: 60 } }));

    // ---- body
    V.body.split('\n').filter(x => x.trim()).forEach(txt => kids.push(P(NP.makeRuns(txt), {
      alignment: d.AlignmentType.JUSTIFIED, spacing: { line: 360 }, indent: { firstLine: 720 }
    })));
    kids.push(P([T('')], { spacing: { after: 40 } }));

    // ---- items table
    const w = [800, 2900, 1500, 1300, 1700];
    const tc = (text, i, o) => new d.TableCell(Object.assign({
      children: [P([T(String(text), { bold: !!(o && o.bold) })], { alignment: (o && o.align) || d.AlignmentType.CENTER })],
      width: { size: w[i], type: d.WidthType.DXA }, borders: allBorders, verticalAlign: d.VerticalAlign.CENTER,
      margins: { top: 40, bottom: 40, left: 90, right: 90 }
    }, (o && o.cell) || {}));
    const rows = [new d.TableRow({
      tableHeader: true, cantSplit: true,
      children: [tc('S.No.', 0, { bold: true }), tc(S.tableHeading || 'Particulars', 1, { bold: true }),
        tc('Qty.(no.)', 2, { bold: true }), tc('Rate (Rs.)', 3, { bold: true }), tc('Amount (Rs.)', 4, { bold: true })]
    })];
    S.items.forEach((it, i) => rows.push(new d.TableRow({
      cantSplit: true,
      children: [tc(i + 1 + '.', 0), tc(it.name, 1, { align: d.AlignmentType.LEFT }), tc(it.qty, 2), tc(it.rate, 3), tc(it.amount, 4)]
    })));
    rows.push(new d.TableRow({
      cantSplit: true,
      children: [
        new d.TableCell({
          children: [P([T('Total', { bold: true })])], columnSpan: 4, borders: allBorders,
          width: { size: w[0] + w[1] + w[2] + w[3], type: d.WidthType.DXA }, verticalAlign: d.VerticalAlign.CENTER,
          margins: { top: 60, bottom: 60, left: 90, right: 90 }
        }),
        tc(V.total.toFixed(2), 4, { bold: true })
      ]
    }));
    kids.push(new d.Table({
      width: { size: w.reduce((a, b) => a + b, 0), type: d.WidthType.DXA }, columnWidths: w,
      alignment: d.AlignmentType.CENTER, borders: allBorders, rows
    }));

    // ---- amount in words, payment, head of account
    kids.push(P(NP.makeRuns(V.wordsLine), { spacing: { before: 140, after: 140 }, indent: { left: 700 } }));
    if (S.includePayment && V.paymentLine)
      kids.push(P(NP.makeRuns(V.paymentLine), { indent: { firstLine: 720 }, spacing: { after: 80 } }));
    kids.push(P([T('The expenditure shall be debited from the following Head of Account.')], { indent: { firstLine: 720 }, spacing: { after: 120 } }));
    V.headQuoted.forEach(l => kids.push(P([T(l, { bold: true })], { indent: { left: 1700 } })));

    // ---- signature
    kids.push(P([T('PRINCIPAL', { bold: true })], { alignment: d.AlignmentType.CENTER, indent: { left: 5200 }, spacing: { before: 700 } }));
    kids.push(P([T(S.officeShort || 'Agricultural Polytechnic, Jagtial', { bold: true })], { alignment: d.AlignmentType.CENTER, indent: { left: 5200 } }));

    // ---- To / CC
    kids.push(P([T('To')], { spacing: { before: 520 } }));
    S.toLines.filter(x => x.trim()).concat(S.ccLines.filter(x => x.trim())).forEach(l => kids.push(P([T(l)])));

    const doc = new d.Document({
      creator: 'Note → Proceeding', title: 'Proceeding ' + S.procNo,
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 20160 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        children: kids
      }]
    });
    return d.Packer.toBlob(doc);
  };
})(window.NP = window.NP || {});
