/* Indian-system number → words, and money helpers */
(function (NP) {
  const ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function two(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? ' ' + ONES[n % 10] : ''); }
  function three(n) {
    const h = Math.floor(n / 100), r = n % 100;
    let s = h ? ONES[h] + ' Hundred' : '';
    if (r) s += (s ? ' and ' : '') + two(r);
    return s;
  }

  function intToWords(n) {
    if (n === 0) return 'Zero';
    const crore = Math.floor(n / 1e7), lakh = Math.floor(n / 1e5) % 100,
      thou = Math.floor(n / 1e3) % 100, rem = n % 1000;
    const parts = [];
    if (crore) parts.push(intToWords(crore) + ' Crore');
    if (lakh) parts.push(two(lakh) + ' Lakh');
    if (thou) parts.push(two(thou) + ' Thousand');
    if (rem) {
      // "Two Thousand and Fifty" – insert "and" when a small remainder follows larger units
      parts.push(parts.length && rem < 100 ? 'and ' + three(rem) : three(rem));
    }
    return parts.join(' ');
  }

  /** 2375 -> "Rupees Two Thousand Three Hundred and Seventy Five only" */
  NP.amountToWords = function (amount) {
    const total = Math.round(Number(amount) * 100);
    const rupees = Math.floor(total / 100), paise = total % 100;
    let s = 'Rupees ' + intToWords(rupees);
    if (paise) s += ' and Paise ' + two(paise);
    return s + ' only';
  };

  /** Indian digit grouping: 123456 -> "1,23,456" */
  NP.inr = function (n) {
    const neg = n < 0; n = Math.abs(n);
    const hasPaise = Math.round(n * 100) % 100 !== 0;
    let [i, d] = n.toFixed(2).split('.');
    const last3 = i.slice(-3), rest = i.slice(0, -3);
    i = (rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' : '') + last3;
    return (neg ? '-' : '') + i + (hasPaise ? '.' + d : '');
  };

  NP.parseNum = function (s) {
    if (s == null) return NaN;
    const m = String(s).replace(/,/g, '').match(/-?\d+(\.\d+)?/);
    return m ? parseFloat(m[0]) : NaN;
  };

  NP.ordinal = function (n) {
    const s = ['th', 'st', 'nd', 'rd'], v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  };
})(window.NP = window.NP || {});
