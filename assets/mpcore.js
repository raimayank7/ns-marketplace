/* NSMarket: reconcile a marketplace payout (settlement) file with the seller's own sales, for marketplace.niveshsahayak.in. Pure functions, no screen.
   The real layout of each marketplace's file is NOT known to this tool and changes: it finds columns by their headings where it can, and the user
   can map every column by hand. Two shapes of file are read:
     wide   one row per order (or order line) with separate columns: sale value, returns, commission, shipping fee, other fees, GST on fees, TCS, TDS, net payout
     long   one row per charge: an order id, a description (Principal, Commission, Shipping fee, TCS ...) and ONE amount column; the description decides the kind
   detect(rows)               header row and column map found from the headings
   parse(rows, map, opts)     rows -> lines { order, date, ref, sale, ret, comm, ship, fee, gstfee, tcs, tds, other, net }
   aggregate(lines)           one entry per order id, totals, one entry per settlement reference
   match(orders, docs, biz)   each order against the seller's saved invoices / cash memos by order id (in the number, reference, notes, item names), then by amount and date;
                              returns against credit notes; sales on the marketplace (by the name the seller types) that are not in the file
   tax(totals, opts)          TCS (s.52 CGST) and TDS (s.194-O) as deducted, and as the typed rates would give, to compare by eye
   books(rec)                 the books side of a payout as one journal entry for NSBooks: bank gets the money, fees are an expense, TCS and TDS are receivable */
(function (G) {
  'use strict';
  var BK = (typeof require === 'function' && typeof module !== 'undefined') ? require('./bankcore.js') : G.NSBank;
  function num(v) { return BK.num(v); }
  function r2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
  function key(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
  function s(v) { return String(v == null ? '' : v).trim(); }

  /* the fields a file can give. 'amt' + 'desc' = the long shape; the others = the wide shape. */
  var FIELDS = ['order', 'date', 'ref', 'desc', 'amt', 'sale', 'ret', 'comm', 'ship', 'fee', 'gstfee', 'tcs', 'tds', 'net', 'taxable'];
  var HEAD = [
    ['order', /order\s*(id|no|number|item)|^order$|sub\s*order|^id$|order-id/i],
    ['date', /(order|settlement|payment|posted|transaction|invoice|payout)?\s*(date|dt)\b|date\/time/i],
    ['ref', /utr|settlement\s*(id|ref|no)|payment\s*(id|ref)|reference|payout\s*(id|ref)|batch/i],
    ['desc', /amount[-\s]*(type|description)|transaction\s*type|charge\s*(type|name)|^type$|description|particulars|fee\s*name|^head$/i],
    ['amt', /^(total\s*)?amount$|^amount\s*\(|^value$|^amt$/i],
    ['sale', /sale\s*(amount|value)|selling\s*price|item\s*price|order\s*value|product\s*(price|amount)|gross\s*(sale|amount)|invoice\s*(value|amount)|principal/i],
    ['ret', /refund|return(s|ed)?\s*(amount|value)?$|^returns?$|reversal|cancel/i],
    ['comm', /commission|referral/i],
    ['ship', /shipping|delivery|fulfil|logistic|freight|weight\s*handling|pick\s*(and|&)?\s*pack/i],
    ['gstfee', /(gst|igst|cgst|sgst|tax)\s*(on|@|charged)?\s*(fee|commission|service|charges)|service\s*tax|gst\s*on/i],
    ['fee', /fee|charges?|closing|collection|storage|ads?\b|penalt|other\s*deduction|marketplace/i],
    ['tcs', /\btcs\b|tax\s*collected\s*at\s*source/i],
    ['tds', /\btds\b|194[\s-]*o|withheld|income\s*tax/i],
    ['net', /net\s*(payout|settlement|amount|payable|receivable)|settlement\s*amount|payout\s*amount|amount\s*(settled|paid)|bank\s*settlement|payable\s*to\s*seller/i],
    ['taxable', /taxable\s*(value|amount)?/i]
  ];
  function detect(rows) {
    var best = { header: -1, map: {}, n: 0 };
    for (var r = 0; r < Math.min((rows || []).length, 40); r++) {
      var map = {}, n = 0;
      (rows[r] || []).forEach(function (c, i) {
        c = s(c); if (!c || c.length > 60) return;
        for (var h = 0; h < HEAD.length; h++) {
          var k = HEAD[h][0];
          if (map[k] == null && HEAD[h][1].test(c)) { map[k] = i; n++; break; }
        }
      });
      if (map.order != null && n > best.n) best = { header: r, map: map, n: n };
    }
    return best;
  }

  /* what a description in a long file means */
  function classify(desc) {
    var d = key(desc);
    if (/\btcs\b|tax collected at source|section 52/.test(d)) return 'tcs';
    if (/\btds\b|194[\s-]*o|income tax|withh/.test(d)) return 'tds';
    if (/(gst|igst|cgst|sgst|service tax|tax)\b.*(fee|commission|service|charge|shipping)|(fee|commission|service|shipping).*(gst|igst|cgst|sgst|tax)\b/.test(d)) return 'gstfee';
    if (/refund|return|reversal|cancel|chargeback/.test(d) && !/fee|commission|shipping/.test(d)) return 'ret';
    if (/commission|referral/.test(d)) return 'comm';
    if (/shipping|delivery|fulfil|logistic|freight|weight|pick/.test(d)) return 'ship';
    if (/principal|order|sale|product|item price|selling|income|revenue/.test(d) && !/fee|charge|commission/.test(d)) return 'sale';
    if (/fee|charge|closing|collection|storage|ads|advertis|penalt|service/.test(d)) return 'fee';
    return 'other';
  }

  function cell(r, map, k) { return map[k] == null ? '' : r[map[k]]; }
  /* rows (arrays of cells, header row excluded or included: the header row is skipped when its index is given) -> lines */
  function parse(rows, map, opts) {
    opts = opts || {};
    var out = { lines: [], skipped: 0, mode: map.amt != null && map.desc != null ? 'long' : 'wide', problems: [] };
    if (map.order == null) { out.problems.push('order'); return out; }
    if (out.mode === 'wide' && map.sale == null && map.net == null && map.amt == null) { out.problems.push('amounts'); return out; }
    (rows || []).forEach(function (r, i) {
      if (opts.header != null && i <= opts.header) return;
      if (!r || !r.some(function (c) { return s(c); })) return;
      var id = s(cell(r, map, 'order')); if (!id || /^total|^grand/i.test(id)) { out.skipped++; return; }
      var L = { order: id, date: BK.date(cell(r, map, 'date')), ref: s(cell(r, map, 'ref')), sale: 0, ret: 0, comm: 0, ship: 0, fee: 0, gstfee: 0, tcs: 0, tds: 0, other: 0, net: null, raw: 0 };
      if (out.mode === 'long') {
        var a = num(cell(r, map, 'amt')), k = classify(cell(r, map, 'desc'));
        L.raw = a; L.net = a;
        if (k === 'other') L.other = a; else L[k] = Math.abs(a);
      } else {
        ['sale', 'ret', 'comm', 'ship', 'fee', 'gstfee', 'tcs', 'tds'].forEach(function (f) { L[f] = Math.abs(num(cell(r, map, f))); });
        if (map.net != null && s(cell(r, map, 'net')) !== '') { L.net = num(cell(r, map, 'net')); L.raw = L.net; }
        if (map.taxable != null) L.taxable = Math.abs(num(cell(r, map, 'taxable')));
        if (map.amt != null && map.sale == null && map.net == null) { L.sale = Math.abs(num(cell(r, map, 'amt'))); }
      }
      out.lines.push(L);
    });
    return out;
  }

  /* one entry per order, in file order; computed = what the payout should be from the parts; diff = what the file paid less computed */
  function aggregate(lines) {
    var by = {}, order = [], refs = {}, rorder = [];
    var T = { orders: 0, sale: 0, ret: 0, comm: 0, ship: 0, fee: 0, gstfee: 0, tcs: 0, tds: 0, other: 0, computed: 0, net: 0, hasNet: false, taxable: 0 };
    lines.forEach(function (L) {
      var o = by[L.order];
      if (!o) { o = by[L.order] = { order: L.order, date: L.date, ref: L.ref, sale: 0, ret: 0, comm: 0, ship: 0, fee: 0, gstfee: 0, tcs: 0, tds: 0, other: 0, net: null, taxable: 0, n: 0 }; order.push(L.order); }
      if (!o.date && L.date) o.date = L.date; if (!o.ref && L.ref) o.ref = L.ref;
      ['sale', 'ret', 'comm', 'ship', 'fee', 'gstfee', 'tcs', 'tds', 'other', 'taxable'].forEach(function (f) { o[f] = r2(o[f] + (L[f] || 0)); });
      if (L.net != null) o.net = r2((o.net || 0) + L.net);
      o.n++;
    });
    var list = order.map(function (id) {
      var o = by[id];
      o.computed = r2(o.sale - o.ret - o.comm - o.ship - o.fee - o.gstfee - o.tcs - o.tds + o.other);
      o.paid = o.net == null ? o.computed : o.net;
      o.diff = o.net == null ? 0 : r2(o.net - o.computed);
      T.orders++; ['sale', 'ret', 'comm', 'ship', 'fee', 'gstfee', 'tcs', 'tds', 'other', 'computed', 'taxable'].forEach(function (f) { T[f] = r2(T[f] + o[f]); });
      T.net = r2(T.net + o.paid); if (o.net != null) T.hasNet = true;
      var rk = o.ref || '—'; if (!refs[rk]) { refs[rk] = { ref: o.ref || '', orders: 0, net: 0, date: o.date }; rorder.push(rk); }
      refs[rk].orders++; refs[rk].net = r2(refs[rk].net + o.paid); if (o.date && (!refs[rk].date || o.date > refs[rk].date)) refs[rk].date = o.date;
      return o;
    });
    T.fees = r2(T.comm + T.ship + T.fee + T.gstfee);
    T.diff = r2(T.net - T.computed);
    var dates = list.map(function (o) { return o.date; }).filter(Boolean).sort();
    return { orders: list, totals: T, refs: rorder.map(function (k) { return refs[k]; }), from: dates[0] || '', to: dates[dates.length - 1] || '' };
  }

  /* ---------- against the seller's own records ---------- */
  var SALES = { invoice: 1, cashmemo: 1, pos: 1, pharma: 1, resto: 1 };
  function tokens(d) {
    var txt = [d.no, d.ref, d.po, d.orderNo, d.notes, d.narr, d.buyer && d.buyer.name].concat((d.items || []).map(function (i) { return i.desc; })).join(' ');
    return (txt.toLowerCase().match(/[a-z0-9][a-z0-9\-_\/]{4,}/g) || []);
  }
  function docTotal(d) { return d.total != null && d.total !== '' ? num(d.total) : 0; }
  function days(a, b) { return Math.abs(Math.round((Date.parse(a + 'T00:00:00Z') - Date.parse(b + 'T00:00:00Z')) / 864e5)); }
  function match(orders, docs, biz, opts) {
    opts = opts || {};
    var b = key(biz), win = opts.days == null ? 10 : opts.days, name = key(opts.name);
    var mine = (docs || []).filter(function (d) { return key(d.seller && d.seller.name) === b; });
    var sales = mine.filter(function (d) { return SALES[d.type || 'invoice']; });
    var notes = mine.filter(function (d) { return d.type === 'note' && d.kind !== 'debit' && d.side !== 'purchase'; });
    function index(list) { var ix = {}; list.forEach(function (d) { tokens(d).forEach(function (t) { (ix[t] = ix[t] || []).push(d); }); }); return ix; }
    var sIx = index(sales), nIx = index(notes), used = {}, usedN = {};
    var res = orders.map(function (o) {
      var id = key(o.order).replace(/\s+/g, ''), how = '', hit = null, nt = null;
      if (id.length >= 6) {
        hit = (sIx[id] || []).filter(function (d) { return !used[d.id]; })[0] || null; if (hit) how = 'id';
        nt = (nIx[id] || []).filter(function (d) { return !usedN[d.id]; })[0] || null;
      }
      return { o: o, doc: hit, how: how, note: nt };
    });
    /* by amount and date for what the order id did not find */
    res.forEach(function (x) { if (x.doc) used[x.doc.id] = 1; if (x.note) usedN[x.note.id] = 1; });
    res.forEach(function (x) {
      if (x.doc || !(x.o.sale > 0)) return;
      var best = null;
      sales.forEach(function (d) {
        if (used[d.id]) return; var t = docTotal(d);
        if (Math.abs(t - x.o.sale) > 1.005 || !x.o.date || !d.date || days(d.date, x.o.date) > win) return;
        if (!best || days(d.date, x.o.date) < days(best.date, x.o.date)) best = d;
      });
      if (best) { x.doc = best; x.how = 'amount'; used[best.id] = 1; }
    });
    res.forEach(function (x) {
      if (x.note || !(x.o.ret > 0)) return;
      var best = null;
      notes.forEach(function (d) { if (usedN[d.id]) return; var t = docTotal(d); if (Math.abs(t - x.o.ret) <= 1.005 && (!best || (d.date || '') > (best.date || ''))) best = d; });
      if (best) { x.note = best; usedN[best.id] = 1; }
    });
    /* sales made on this marketplace (the seller types its name as written on his invoices) that are not in the file */
    var from = opts.from || '', to = opts.to || '';
    var missing = !name ? [] : sales.filter(function (d) {
      if (used[d.id] || (from && d.date < from) || (to && d.date > to)) return false;
      return (tokens(d).join(' ') + ' ' + key(d.notes) + ' ' + key(d.ref) + ' ' + key(d.buyer && d.buyer.name)).indexOf(name) >= 0;
    });
    return { rows: res, missing: missing };
  }

  /* TCS (s.52 CGST) and TDS (s.194-O): as deducted in the file and as the typed rates would give. The rates are typed by the user, never assumed. */
  function tax(T, opts) {
    opts = opts || {};
    var g = num(opts.gst), base = T.taxable > 0 ? T.taxable : (g ? r2((T.sale - T.ret) * 100 / (100 + g)) : r2(T.sale - T.ret));
    var out = { base: r2(base), gross: r2(T.sale - T.ret), tcs: T.tcs, tds: T.tds, tcsRate: num(opts.tcsRate), tdsRate: num(opts.tdsRate) };
    out.tcsExpected = out.tcsRate ? r2(base * out.tcsRate / 100) : null;
    out.tdsExpected = out.tdsRate ? r2(out.gross * out.tdsRate / 100) : null;
    out.tcsDiff = out.tcsExpected == null ? null : r2(T.tcs - out.tcsExpected);
    out.tdsDiff = out.tdsExpected == null ? null : r2(T.tds - out.tdsExpected);
    return out;
  }

  /* books: the marketplace is a party (receivable); the payout received, the fees (an expense), TCS and TDS (receivable) clear what it owes.
     Returns are in the seller's own credit notes, so the party owes sales less returns = net + fees + TCS + TDS. */
  function books(rec) {
    var t = rec.totals || {}, P = String(rec.party || '').trim(), L = [];
    var fees = r2(num(t.fees)), tcs = r2(num(t.tcs)), tds = r2(num(t.tds)), recv = r2(num(rec.received));
    if (!P || !(recv > 0 || fees > 0 || tcs > 0 || tds > 0)) return [];
    function line(a, dr, cr, x) { if (dr || cr) L.push(Object.assign({ a: a, dr: r2(dr), cr: r2(cr) }, x || {})); }
    var mode = rec.mode === 'cash' ? 'cash' : 'bank';
    line(mode, recv, 0); line('exp:marketplace fees and charges', fees, 0); line('tcs-receivable', tcs, 0); line('tds-receivable', tds, 0);
    line('party', 0, r2(recv + fees + tcs + tds), { party: P, side: 'recv' });
    return [{ type: 'jv', no: rec.no || '', narr: 'Marketplace payout ' + (rec.name || '') + ' ' + (rec.from || '') + ' to ' + (rec.to || ''), lines: L, date: rec.payDate || rec.to }];
  }

  /* a made-up example file in both shapes, for the user to see the columns (not any marketplace's real layout) */
  function sample() {
    return {
      wide: [['Order ID', 'Order date', 'Settlement ref (UTR)', 'Sale value', 'Refund', 'Commission', 'Shipping fee', 'Other fees', 'GST on fees', 'TCS', 'TDS', 'Net payout'],
        ['ORD-1001', '2026-10-02', 'UTR100200', 1000, 0, 100, 60, 10, 30.6, 5, 1, 793.4], ['ORD-1002', '2026-10-03', 'UTR100200', 2500, 0, 250, 80, 20, 63, 12.5, 2.5, 2072],
        ['ORD-1003', '2026-10-04', 'UTR100200', 1500, 1500, 150, 60, 0, 37.8, 7.5, 1.5, -256.8]],
      long: [['Order ID', 'Posted date', 'Settlement ID', 'Charge type', 'Amount'],
        ['ORD-2001', '2026-10-02', 'SET900', 'Item principal', 1000], ['ORD-2001', '2026-10-02', 'SET900', 'Commission', -100], ['ORD-2001', '2026-10-02', 'SET900', 'Shipping fee', -60], ['ORD-2001', '2026-10-02', 'SET900', 'GST on fees', -28.8],
        ['ORD-2001', '2026-10-02', 'SET900', 'TCS', -5], ['ORD-2001', '2026-10-02', 'SET900', 'TDS 194-O', -1], ['ORD-2002', '2026-10-03', 'SET900', 'Item principal', 2000], ['ORD-2002', '2026-10-03', 'SET900', 'Commission', -200]]
    };
  }

  var API = { num: num, r2: r2, FIELDS: FIELDS, detect: detect, classify: classify, parse: parse, aggregate: aggregate, match: match, tax: tax, books: books, sample: sample, tokens: tokens };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSMarket = API;
})(this);
