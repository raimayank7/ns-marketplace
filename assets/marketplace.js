/* marketplace.niveshsahayak.in: reconcile a marketplace payout (settlement) file with your own sales.
   The file (CSV or Excel) is read in this browser only (NSMarket in mpcore.js). The layout of each marketplace's file is not known to this tool:
   it finds the columns by their headings and every column can be changed by hand. Each order is matched with your saved invoices by order id,
   then by amount and date; returns, fees, TCS and TDS are totalled; what the file paid is compared with what its own parts add up to, and
   with what reached your bank. The payout can be added to the books of the free NSBiz (record type 'mpay', books side in d.books). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, num = NS.num, r2 = NS.r2, money = NS.money, M = NSMarket, t = NSB.t;
  var LIB = { xl: ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'sha384-vtjasyidUo0kW94K5MXDXntzOJpQgBKXmE7e2Ga4LG0skTTLeBi97eFAXsqewJjw'] };
  var cur = 'home', ST = null, OPT = load(), filter = 'all', limit = 200, saveId = '';
  /* ST = { name, sheets: [[name, rows]], sheet, rows, header, map, parse, ag, mt } */
  function load() { var o = {}; try { o = JSON.parse(localStorage.getItem('ns-marketplace-opts') || '{}'); } catch (e) {} return Object.assign({ name: '', tcsRate: '', tdsRate: '', gst: '', days: 10, party: '', received: '', mode: 'bank', payDate: '' }, o); }
  function keep() { try { localStorage.setItem('ns-marketplace-opts', JSON.stringify(OPT)); } catch (e) {} }
  function D() { return NSB.data().doc; }
  function mine() { return NSB.mine(); }
  function route() { cur = location.hash.slice(1) === 'saved' ? 'saved' : 'home'; show(true); }
  function show(focus) {
    NSB.navMark(cur);
    return NSB.load().then(function () {
      var pane = $('[data-pane]'); NSB.modeLine();
      if (!NSB.biz()) { pane.innerHTML = NSB.noBiz(); return; }
      if (ST && ST.rows) work();   /* invoices saved since the file was opened count too */
      pane.innerHTML = NSB.bizBar() + (cur === 'saved' ? savedHtml() : homeHtml());
      wire(pane); if (focus) NSB.focusH2(pane);
    });
  }

  /* ---------- the work ---------- */
  function work() {
    var p = M.parse(ST.rows, ST.map, { header: ST.header });
    ST.parse = p;
    if (p.problems.length) { ST.ag = null; ST.mt = null; return; }
    ST.ag = M.aggregate(p.lines);
    ST.mt = M.match(ST.ag.orders, D(), NSB.bizName(), { name: OPT.name, days: num(OPT.days) || 10, from: ST.ag.from, to: ST.ag.to });
    ST.tx = M.tax(ST.ag.totals, { gst: OPT.gst, tcsRate: OPT.tcsRate, tdsRate: OPT.tdsRate });
  }
  function setSheet(i) {
    var sh = ST.sheets[i]; ST.sheet = i; ST.rows = sh[1];
    var d = M.detect(sh[1]); ST.header = d.header; ST.map = d.map; work();
  }
  function readFile(file) {
    var xl = /\.xlsx?$/i.test(file.name);
    var get = xl ? NS.loadJs(LIB.xl[0], LIB.xl[1]).then(function () { return file.arrayBuffer(); }).then(function (buf) {
      var wb = window.XLSX.read(buf, { type: 'array', cellDates: false });
      return wb.SheetNames.map(function (n) { return [n, window.XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' })]; });
    }) : file.text().then(function (x) { return [['CSV', NSBank.csvRows(x)]]; });
    return get.then(function (sheets) { startWith(sheets.filter(function (s) { return s[1].length; }), file.name); }).catch(function () { NS.toast(t('mk.readfail'), 'bad'); });
  }
  function startWith(sheets, name) {
    if (!sheets.length) { NS.toast(t('mk.readfail'), 'bad'); return; }
    ST = { name: name, sheets: sheets }; saveId = ''; filter = 'all'; limit = 200;
    /* the first sheet that has an order column, else the first */
    var pick = 0; for (var i = 0; i < sheets.length; i++) { if (M.detect(sheets[i][1]).header >= 0) { pick = i; break; } }
    setSheet(pick); show();
    if (ST.ag) NS.toast(t('mk.read', { n: ST.ag.totals.orders }));
  }

  /* ---------- screen ---------- */
  var FL = ['order', 'date', 'ref', 'desc', 'amt', 'sale', 'ret', 'comm', 'ship', 'fee', 'gstfee', 'tcs', 'tds', 'net', 'taxable'];
  function colsHtml() {
    var n = Math.max.apply(null, ST.rows.slice(0, 50).map(function (r) { return r.length; }).concat([1])), hdr = ST.rows[ST.header] || [];
    var opt = function (k) { return '<option value="">—</option>' + Array.apply(null, Array(n)).map(function (_, i) { return '<option value="' + i + '"' + (ST.map[k] === i ? ' selected' : '') + '>' + esc((i + 1) + ': ' + String(hdr[i] == null || hdr[i] === '' ? '' : hdr[i]).slice(0, 28)) + '</option>'; }).join(''); };
    var sheets = ST.sheets.length > 1 ? '<div class="f"><label for="mk-sheet">' + esc(t('mk.sheet')) + '</label><select id="mk-sheet">' + ST.sheets.map(function (s, i) { return '<option value="' + i + '"' + (i === ST.sheet ? ' selected' : '') + '>' + esc(s[0]) + '</option>'; }).join('') + '</select></div>' : '';
    var hr = '<div class="f"><label for="mk-hdr">' + esc(t('mk.hdrrow')) + '</label><input id="mk-hdr" type="number" min="1" max="60" value="' + (ST.header + 1) + '"></div>';
    return '<p class="hint">' + esc(t('mk.colsh')) + '</p><div class="fg br-map">' + sheets + hr + FL.map(function (k) { return '<div class="f"><label for="mk-c-' + k + '">' + esc(t('mk.c_' + k)) + '</label><select id="mk-c-' + k + '" data-col="' + k + '">' + opt(k) + '</select></div>'; }).join('') + '</div>';
  }
  function settingsHtml() {
    var f = function (id, lbl, v, extra, hint) { return '<div class="f"><label for="mk-' + id + '">' + esc(t(lbl)) + '</label><input id="mk-' + id + '" data-opt="' + id + '" ' + (extra || '') + ' value="' + esc(v) + '">' + (hint ? '<small class="hint">' + esc(t(hint)) + '</small>' : '') + '</div>'; };
    return '<div class="fg">' + f('name', 'mk.mname', OPT.name, 'dir="auto" placeholder="Amazon, Flipkart, Meesho…"', 'mk.mnameh') + f('days', 'mk.days', OPT.days, 'type="number" min="0" max="60"') +
      f('tcsRate', 'mk.tcsrate', OPT.tcsRate, 'inputmode="decimal"', 'mk.tcsrateh') + f('tdsRate', 'mk.tdsrate', OPT.tdsRate, 'inputmode="decimal"', 'mk.tdsrateh') + f('gst', 'mk.gst', OPT.gst, 'inputmode="decimal"', 'mk.gsth') + '</div>';
  }
  function homeHtml() {
    var head = '<div class="pane-hd"><h2>' + esc(t('mk.title')) + '</h2><div class="quick"><label class="btn primary" for="mk-file">' + esc(t(ST ? 'mk.other' : 'mk.open')) + '</label>' +
      '<input type="file" id="mk-file" class="sr" accept=".csv,.txt,.xls,.xlsx,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"><button type="button" class="btn" data-sample>' + esc(t('mk.sample')) + '</button></div></div>';
    if (!ST) return head + '<div class="card"><h3>' + esc(t('mk.start')) + '</h3><p>' + esc(t('mk.starth')) + '</p><ol class="br-steps"><li>' + esc(t('mk.s1')) + '</li><li>' + esc(t('mk.s2')) + '</li><li>' + esc(t('mk.s3')) + '</li></ol><p class="sb-note">' + esc(t('mk.layout')) + '</p></div>';
    var P = ST.parse, info = '<p class="sb-note">' + esc(t('mk.from', { f: ST.name, m: t('mk.mode.' + (P.mode || 'wide')) })) + '</p><details class="br-cols"' + (ST.ag ? '' : ' open') + '><summary>' + esc(t('mk.cols')) + '</summary>' + colsHtml() + '</details>';
    if (!ST.ag) return head + info + '<div class="card"><p class="sb-note bad">' + esc(t('mk.prob.' + P.problems[0])) + '</p></div>';
    var A = ST.ag, T = A.totals, X = ST.tx, rows = ST.mt.rows;
    var nMiss = rows.filter(function (r) { return !r.doc && r.o.sale > 0; }).length, nRet = rows.filter(function (r) { return r.o.ret > 0; }).length, nNoCn = rows.filter(function (r) { return r.o.ret > 0 && !r.note; }).length, nDiff = rows.filter(function (r) { return Math.abs(r.o.diff) > 0.005; }).length;
    var recv = OPT.received === '' ? null : r2(num(OPT.received)), short = recv == null ? null : r2(T.net - recv);
    var tiles = '<div class="tiles"><div class="tile"><span>' + esc(t('mk.orders')) + '</span><b>' + T.orders + '</b></div><div class="tile"><span>' + esc(t('mk.sales')) + '</span><b>' + money(T.sale) + '</b></div>' +
      '<div class="tile' + (T.ret ? ' warn' : '') + '"><span>' + esc(t('mk.returns')) + '</span><b>' + money(T.ret) + '</b></div><div class="tile"><span>' + esc(t('mk.fees')) + '</span><b>' + money(T.fees) + '</b></div>' +
      '<div class="tile"><span>' + esc(t('mk.tcs')) + ' / ' + esc(t('mk.tds')) + '</span><b>' + money(T.tcs + T.tds) + '</b></div><div class="tile ok"><span>' + esc(t('mk.payout')) + '</span><b>' + money(T.net) + '</b></div></div>';
    var flags = '<ul class="mk-flags">' + (nMiss ? '<li class="warn">' + esc(t('mk.f.miss', { n: nMiss })) + '</li>' : '') + (ST.mt.missing.length ? '<li class="warn">' + esc(t('mk.f.nofile', { n: ST.mt.missing.length })) + '</li>' : '') +
      (nNoCn ? '<li class="warn">' + esc(t('mk.f.nocn', { n: nNoCn })) + '</li>' : '') + (nDiff ? '<li class="warn">' + esc(t('mk.f.diff', { n: nDiff, v: money(Math.abs(T.diff)) })) + '</li>' : '') +
      (short != null && Math.abs(short) > 0.005 ? '<li class="warn">' + esc(t(short > 0 ? 'mk.f.short' : 'mk.f.more', { v: money(Math.abs(short)) })) + '</li>' : short != null ? '<li class="ok">' + esc(t('mk.f.ok')) + '</li>' : '') + '</ul>';
    var taxh = '<h3>' + esc(t('mk.taxh')) + '</h3><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l"></th><th class="n">' + esc(t('mk.deducted')) + '</th><th class="n">' + esc(t('mk.expected')) + '</th><th class="n">' + esc(t('mk.diff')) + '</th></tr></thead><tbody>' +
      '<tr><td class="l">' + esc(t('mk.tcsl')) + '</td><td class="n">' + money(X.tcs, false) + '</td><td class="n">' + (X.tcsExpected == null ? '—' : money(X.tcsExpected, false)) + '</td><td class="n">' + (X.tcsDiff == null ? '—' : money(X.tcsDiff, false)) + '</td></tr>' +
      '<tr><td class="l">' + esc(t('mk.tdsl')) + '</td><td class="n">' + money(X.tds, false) + '</td><td class="n">' + (X.tdsExpected == null ? '—' : money(X.tdsExpected, false)) + '</td><td class="n">' + (X.tdsDiff == null ? '—' : money(X.tdsDiff, false)) + '</td></tr></tbody></table></div><p class="hint">' + esc(t('mk.taxnote')) + '</p>';
    var recvBox = '<div class="fg"><div class="f"><label for="mk-received">' + esc(t('mk.received')) + '</label><input id="mk-received" data-opt="received" inputmode="decimal" value="' + esc(OPT.received) + '"><small class="hint">' + esc(t('mk.receivedh')) + '</small></div>' +
      '<div class="f"><label for="mk-payDate">' + esc(t('mk.paydate')) + '</label><input id="mk-payDate" data-opt="payDate" type="date" value="' + esc(OPT.payDate || A.to) + '"></div></div>';
    var refs = '<details class="br-pairs"><summary>' + esc(t('mk.refs')) + ' (' + A.refs.length + ')</summary><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">' + esc(t('mk.c_ref')) + '</th><th>' + esc(t('mk.c_date')) + '</th><th class="n">' + esc(t('mk.orders')) + '</th><th class="n">' + esc(t('mk.payout')) + '</th></tr></thead><tbody>' +
      A.refs.map(function (x) { return '<tr><td class="l mono">' + esc(x.ref || '—') + '</td><td>' + (x.date ? NS.fmtDate(x.date) : '') + '</td><td class="n">' + x.orders + '</td><td class="n">' + money(x.net, false) + '</td></tr>'; }).join('') + '</tbody></table></div><p class="hint">' + esc(t('mk.refsh')) + '</p></details>';
    var chips = [['all', 'mk.fl.all', rows.length], ['miss', 'mk.fl.miss', nMiss], ['ret', 'mk.fl.ret', nRet], ['diff', 'mk.fl.diff', nDiff]].map(function (c) { return '<button type="button" class="btn sm' + (filter === c[0] ? ' primary' : '') + '" data-filter="' + c[0] + '">' + esc(t(c[1])) + ' (' + c[2] + ')</button>'; }).join(' ');
    var shown = rows.filter(function (r) { return filter === 'miss' ? !r.doc && r.o.sale > 0 : filter === 'ret' ? r.o.ret > 0 : filter === 'diff' ? Math.abs(r.o.diff) > 0.005 : true; });
    var tbl = '<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">' + esc(t('mk.c_order')) + '</th><th>' + esc(t('mk.c_date')) + '</th><th class="n">' + esc(t('mk.sales')) + '</th><th class="n">' + esc(t('mk.returns')) + '</th><th class="n">' + esc(t('mk.fees')) + '</th><th class="n">' + esc(t('mk.tcs')) + '+' + esc(t('mk.tds')) + '</th><th class="n">' + esc(t('mk.payout')) + '</th><th class="n">' + esc(t('mk.diff')) + '</th><th class="l">' + esc(t('mk.yours')) + '</th></tr></thead><tbody>' +
      shown.slice(0, limit).map(function (r) {
        var o = r.o, fees = r2(o.comm + o.ship + o.fee + o.gstfee), mt = r.doc ? esc(r.doc.no || '') + ' <span class="soft">' + esc(t('mk.how.' + r.how)) + '</span>' : (o.sale > 0 ? '<b>' + esc(t('mk.nosale')) + '</b>' : '');
        if (o.ret > 0) mt += (mt ? '<br>' : '') + (r.note ? esc(r.note.no || '') + ' <span class="soft">' + esc(t('mk.cn')) + '</span>' : '<b>' + esc(t('mk.nocn')) + '</b>');
        return '<tr' + (!r.doc && o.sale > 0 || Math.abs(o.diff) > 0.005 ? ' class="warn"' : '') + '><td class="l mono">' + esc(o.order) + '</td><td>' + (o.date ? NS.fmtDate(o.date) : '') + '</td><td class="n">' + money(o.sale, false) + '</td><td class="n">' + (o.ret ? money(o.ret, false) : '') + '</td><td class="n">' + money(fees, false) + '</td><td class="n">' + money(o.tcs + o.tds, false) + '</td><td class="n">' + money(o.paid, false) + '</td><td class="n">' + (Math.abs(o.diff) > 0.005 ? money(o.diff, false) : '') + '</td><td class="l">' + mt + '</td></tr>';
      }).join('') + '</tbody></table></div>' + (shown.length > limit ? '<p><button type="button" class="btn" data-more>' + esc(t('mk.more', { n: shown.length - limit })) + '</button></p>' : '') + (!shown.length ? '<p class="sb-empty">' + esc(t('mk.none')) + '</p>' : '');
    var miss = ST.mt.missing.length ? '<h3>' + esc(t('mk.nofile')) + ' (' + ST.mt.missing.length + ')</h3><div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">' + esc(t('mk.inv')) + '</th><th>' + esc(t('mk.c_date')) + '</th><th class="l">' + esc(t('mk.cust')) + '</th><th class="n">' + esc(t('mk.amount')) + '</th></tr></thead><tbody>' +
      ST.mt.missing.map(function (d) { return '<tr><td class="l">' + esc(d.no || '') + '</td><td>' + NS.fmtDate(d.date) + '</td><td class="l">' + esc(d.buyer && d.buyer.name || '') + '</td><td class="n">' + money(num(d.total), false) + '</td></tr>'; }).join('') + '</tbody></table></div>' : '';
    return head + info + '<div class="card"><h3>' + esc(t('mk.settings')) + '</h3>' + settingsHtml() + '</div>' + tiles + flags + '<h3>' + esc(t('mk.recvh')) + '</h3>' + recvBox + taxh + refs + '<h3>' + esc(t('mk.orderlist')) + '</h3><p>' + chips + '</p>' + tbl + miss +
      '<div class="card"><h3>' + esc(t('mk.booksh')) + '</h3><p class="hint">' + esc(t('mk.booksp')) + '</p><div class="fg"><div class="f"><label for="mk-party">' + esc(t('mk.party')) + '</label><input id="mk-party" data-opt="party" dir="auto" list="mk-parties" value="' + esc(OPT.party || OPT.name) + '"><datalist id="mk-parties">' +
      NSB.data().party.map(function (p) { return '<option value="' + esc(p.name) + '">'; }).join('') + '</datalist></div><div class="f"><label for="mk-mode">' + esc(t('mk.mode')) + '</label><select id="mk-mode" data-opt="mode"><option value="bank"' + (OPT.mode === 'bank' ? ' selected' : '') + '>' + esc(t('pm.bank')) + '</option><option value="cash"' + (OPT.mode === 'cash' ? ' selected' : '') + '>' + esc(t('pm.cash')) + '</option></select></div></div></div>' +
      '<div class="btns"><button type="button" class="btn primary" data-save>' + esc(t('mk.save')) + '</button><button type="button" class="btn" data-addbooks>' + esc(t('mk.addbooks')) + '</button><button type="button" class="btn" data-xl>' + esc(t('mk.xl')) + '</button><button type="button" class="btn" data-print>' + esc(t('sb.print')) + '</button><button type="button" class="btn" data-pdf>' + esc(t('sb.pdf')) + '</button></div>' +
      '<div class="sb-print">' + sheet() + '</div>';
  }

  /* English print sheet */
  function sheet() {
    var b = NSB.biz() || {}, T = ST.ag.totals, X = ST.tx, rows = ST.mt.rows;
    function tr(l, v, s) { return '<tr' + (s ? ' class="b"' : '') + '><td>' + esc(l) + '</td><td class="n">' + money(v, false) + '</td></tr>'; }
    return '<div class="stsheet" data-sheet><div class="st-hd"><div><h3>' + esc(b.name || '') + '</h3></div><div class="r"><b>Marketplace payout reconciliation</b><div class="soft">' + esc(OPT.name || '') + ' · ' + NS.fmtDate(ST.ag.from) + ' to ' + NS.fmtDate(ST.ag.to) + ' · file: ' + esc(ST.name) + '</div></div></div>' +
      '<table class="tbl"><tbody>' + tr('Orders in the file (count ' + T.orders + ')', T.sale, 1) + tr('Less: returns and refunds', T.ret) + tr('Less: commission', T.comm) + tr('Less: shipping and fulfilment fees', T.ship) + tr('Less: other fees', T.fee) + tr('Less: GST on fees', T.gstfee) + tr('Less: TCS (section 52 CGST)', T.tcs) + tr('Less: TDS (section 194-O)', T.tds) +
      (T.other ? tr('Other adjustments (added)', T.other) : '') + tr('Payout worked out from the parts', T.computed, 1) + (T.hasNet ? tr('Payout as the file shows', T.net, 1) + tr('Difference', T.diff, 1) : '') + (OPT.received !== '' ? tr('Received in the bank (typed)', num(OPT.received), 1) + tr('Payout not received', r2(T.net - num(OPT.received)), 1) : '') + '</tbody></table>' +
      '<h4>Not found in your sales (' + rows.filter(function (r) { return !r.doc && r.o.sale > 0; }).length + ')</h4><table class="tbl"><thead><tr><th>Order</th><th>Date</th><th class="n">Sale</th></tr></thead><tbody>' + rows.filter(function (r) { return !r.doc && r.o.sale > 0; }).slice(0, 80).map(function (r) { return '<tr><td>' + esc(r.o.order) + '</td><td>' + (r.o.date ? NS.fmtDate(r.o.date) : '') + '</td><td class="n">' + money(r.o.sale, false) + '</td></tr>'; }).join('') + '</tbody></table>' +
      '<h4>Returns without a credit note (' + rows.filter(function (r) { return r.o.ret > 0 && !r.note; }).length + ')</h4><table class="tbl"><thead><tr><th>Order</th><th>Date</th><th class="n">Returned</th></tr></thead><tbody>' + rows.filter(function (r) { return r.o.ret > 0 && !r.note; }).slice(0, 80).map(function (r) { return '<tr><td>' + esc(r.o.order) + '</td><td>' + (r.o.date ? NS.fmtDate(r.o.date) : '') + '</td><td class="n">' + money(r.o.ret, false) + '</td></tr>'; }).join('') + '</tbody></table>' +
      (X.tcsExpected != null || X.tdsExpected != null ? '<h4>Tax deducted, for checking</h4><table class="tbl"><tbody>' + tr('TCS deducted', X.tcs) + (X.tcsExpected != null ? tr('TCS at ' + X.tcsRate + '% of ' + money(X.base), X.tcsExpected) : '') + tr('TDS deducted', X.tds) + (X.tdsExpected != null ? tr('TDS at ' + X.tdsRate + '% of ' + money(X.gross), X.tdsExpected) : '') + '</tbody></table>' : '') +
      '<p class="small">The file layout is read as mapped by the user; check against the marketplace’s own report.</p>' + NS.creditLine('st-ft') + '</div>';
  }

  /* ---------- saving and the books ---------- */
  function nextNo(date) {
    var fy = NS.fy(date), short = fy.slice(2, 4) + '-' + fy.slice(5), max = 0;
    mine().forEach(function (d) { if (d.type !== 'mpay' || NS.fy(d.date) !== fy) return; var m = String(d.no || '').match(/(\d+)\s*$/); if (m) max = Math.max(max, +m[1]); });
    return 'MP/' + short + '/' + String(max + 1).padStart(3, '0');
  }
  function record(withBooks) {
    var T = ST.ag.totals, existing = saveId && D().filter(function (d) { return d.id === saveId; })[0];
    var date = OPT.payDate || ST.ag.to, rec = existing || { type: 'mpay', app: 1 };
    Object.assign(rec, { seller: NSB.sellerCopy(), date: date, no: rec.no || nextNo(date), name: OPT.name, party: String(OPT.party || OPT.name || '').trim(), from: ST.ag.from, to: ST.ag.to, payDate: date, mode: OPT.mode, received: OPT.received, file: ST.name,
      totals: T, refs: ST.ag.refs, tax: { tcs: T.tcs, tds: T.tds }, total: T.net, buyer: { name: '' } });
    rec.books = withBooks ? M.books(rec) : (existing && existing.books && existing.books.length && !withBooks ? existing.books : []);
    return rec;
  }
  function save(withBooks) {
    if (withBooks) {
      var P = String(OPT.party || OPT.name || '').trim();
      if (!P) { NS.toast(t('mk.v.party'), 'bad'); return; }
      if (OPT.received === '' || !(num(OPT.received) > 0)) { NS.toast(t('mk.v.recv'), 'bad'); return; }
      if (!NS.okSync(t('mk.c.books'))) return;
    }
    var rec = record(!!withBooks);
    NSStore.put('doc', rec).then(function (x) { saveId = x && x.id || rec.id; NS.toast(t(withBooks ? 'mk.booked' : 'sb.saved')); show(); });
  }
  function fd(d) { return d ? NS.fmtDate(d) : ''; }
  function xl() {
    var A = ST.ag, rows = ST.mt.rows, T = A.totals;
    NS.xlsx('marketplace-reconciliation-' + (A.to || NS.today()) + '.xlsx', [
      ['Summary', [['Item', 'Amount'], ['Orders', T.orders], ['Sales', T.sale], ['Returns', T.ret], ['Commission', T.comm], ['Shipping and fulfilment fees', T.ship], ['Other fees', T.fee], ['GST on fees', T.gstfee], ['TCS', T.tcs], ['TDS', T.tds], ['Payout worked out', T.computed], ['Payout as in the file', T.net], ['Difference', T.diff], ['Received (typed)', OPT.received === '' ? '' : num(OPT.received)]]],
      ['Orders', [['Order', 'Date', 'Settlement ref', 'Sale', 'Returns', 'Commission', 'Shipping', 'Other fees', 'GST on fees', 'TCS', 'TDS', 'Payout (file)', 'Payout (worked out)', 'Difference', 'Your invoice', 'Matched by', 'Credit note']].concat(rows.map(function (r) { var o = r.o; return [o.order, fd(o.date), o.ref, o.sale, o.ret, o.comm, o.ship, o.fee, o.gstfee, o.tcs, o.tds, o.paid, o.computed, o.diff, r.doc ? r.doc.no : '', r.how, r.note ? r.note.no : '']; }))],
      ['Not in your sales', [['Order', 'Date', 'Sale']].concat(rows.filter(function (r) { return !r.doc && r.o.sale > 0; }).map(function (r) { return [r.o.order, fd(r.o.date), r.o.sale]; }))],
      ['Returns, no credit note', [['Order', 'Date', 'Returned']].concat(rows.filter(function (r) { return r.o.ret > 0 && !r.note; }).map(function (r) { return [r.o.order, fd(r.o.date), r.o.ret]; }))],
      ['Sold, not in the file', [['Invoice', 'Date', 'Customer', 'Amount']].concat(ST.mt.missing.map(function (d) { return [d.no, fd(d.date), d.buyer && d.buyer.name || '', num(d.total)]; }))],
      ['Settlements', [['Reference', 'Date', 'Orders', 'Payout']].concat(A.refs.map(function (x) { return [x.ref, fd(x.date), x.orders, x.net]; }))]].map(function (s) { return { name: s[0], rows: s[1] }; }));
  }
  function sampleFile(kind) {
    NS.xlsx('payout-example-' + kind + '.xlsx', [{ name: 'Example', rows: M.sample()[kind] }, { name: 'Read me', rows: [['Read me'], ['This is a made-up example. It is not the layout of any marketplace. Your file can have other column headings: open it here and choose which column is which.']] }]);
  }

  function savedHtml() {
    var list = mine().filter(function (d) { return d.type === 'mpay'; }).sort(function (a, b) { return (b.date || '').localeCompare(a.date || ''); });
    return '<div class="pane-hd"><h2>' + esc(t('mk.saved')) + '</h2></div>' + (list.length ? '<div class="tbl-wrap"><table class="tbl"><thead><tr><th class="l">' + esc(t('mk.number')) + '</th><th>' + esc(t('mk.period')) + '</th><th class="l">' + esc(t('mk.mname')) + '</th><th class="n">' + esc(t('mk.payout')) + '</th><th class="n">' + esc(t('mk.fees')) + '</th><th>' + esc(t('mk.inbooks')) + '</th><th></th></tr></thead><tbody>' +
      list.map(function (d) { var T = d.totals || {}; return '<tr><td class="l mono">' + esc(d.no) + '</td><td>' + NS.fmtDate(d.from) + ' – ' + NS.fmtDate(d.to) + '</td><td class="l">' + esc(d.name || '') + '</td><td class="n">' + money(num(T.net), false) + '</td><td class="n">' + money(num(T.fees) + num(T.tcs) + num(T.tds), false) + '</td><td>' + (d.books && d.books.length ? '✓' : '') + '</td><td><button type="button" class="btn sm danger" data-del="' + esc(d.id) + '">' + esc(t('h.del')) + '</button></td></tr>'; }).join('') + '</tbody></table></div><p class="hint">' + esc(t('mk.savedh')) + '</p>' : '<div class="card"><p>' + esc(t('mk.nosaved')) + '</p></div>');
  }

  function wire(pane) {
    NSB.wireBiz(pane, show);
    var fi = $('#mk-file', pane); if (fi) fi.addEventListener('change', function () { if (fi.files[0]) readFile(fi.files[0]); fi.value = ''; });
    var sm = $('[data-sample]', pane); if (sm) sm.addEventListener('click', function () { (NS.ask ? NS.ask(t('mk.samplec')) : Promise.resolve(confirm(t('mk.samplec')))).then(function (y) { sampleFile(y ? 'wide' : 'long'); }); });
    $$('[data-opt]', pane).forEach(function (e) { e.addEventListener('change', function () { OPT[e.getAttribute('data-opt')] = e.value; keep(); if (ST && ST.rows) work(); show(); }); });
    $$('[data-col]', pane).forEach(function (sel) { sel.addEventListener('change', function () { var m = Object.assign({}, ST.map), k = sel.getAttribute('data-col'); if (sel.value === '') delete m[k]; else m[k] = +sel.value; ST.map = m; work(); show(); }); });
    var sh = $('#mk-sheet', pane); if (sh) sh.addEventListener('change', function () { setSheet(+sh.value); show(); });
    var hr = $('#mk-hdr', pane); if (hr) hr.addEventListener('change', function () { ST.header = Math.max(0, (+hr.value || 1) - 1); work(); show(); });
    $$('[data-filter]', pane).forEach(function (b) { b.addEventListener('click', function () { filter = b.getAttribute('data-filter'); limit = 200; show(); }); });
    var mo = $('[data-more]', pane); if (mo) mo.addEventListener('click', function () { limit += 400; show(); });
    var sv = $('[data-save]', pane); if (sv) sv.addEventListener('click', function () { save(false); });
    var ab = $('[data-addbooks]', pane); if (ab) ab.addEventListener('click', function () { save(true); });
    var x = $('[data-xl]', pane); if (x) x.addEventListener('click', xl);
    var pr = $('[data-print]', pane), pdf = $('[data-pdf]', pane);
    if (pr) pr.addEventListener('click', NSB.printSheet);
    if (pdf) pdf.addEventListener('click', function () { NSB.shareSheet($('[data-sheet]', pane), NSB.bizName() + ' marketplace payout ' + (ST.ag.to || '')); });
    $$('[data-del]', pane).forEach(function (b) { b.addEventListener('click', function () { var d = D().filter(function (x) { return x.id === b.getAttribute('data-del'); })[0]; if (d && NS.okSync(t('sb.cdel', { x: d.no }))) NSStore.del('doc', d.id).then(function () { NS.toast(t('sb.deleted')); show(); }); }); });
  }

  function init() {
    if (!$('#marketplace')) return;
    window.addEventListener('hashchange', route);
    document.addEventListener('ns-lang', function () { show(); });
    NSStore.ready().then(route).then(function () { NSStore.persist().catch(function () {}); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
