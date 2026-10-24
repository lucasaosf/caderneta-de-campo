/* Lista de Campo — gerador da planilha de análises (.xlsx e .csv)
   Identidade visual: Pesquisa & Soluções Ambientais · Lucas Aøsf
   Funciona no navegador e no Node (para testes). Sem bibliotecas externas. */
(function (root) {
  'use strict';

  /* ===================== utilidades ===================== */
  var enc = typeof TextEncoder !== 'undefined' ? new TextEncoder() : null;
  function u8(s) { return enc.encode(s); }
  var CRC = (function () { var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zip(files) {
    var d = new Date(), dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    var parts = [], central = [], off = 0, cd = 0;
    files.forEach(function (f) {
      var name = u8(f.name), data = typeof f.data === 'string' ? u8(f.data) : f.data, crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, tm, true); h.setUint16(12, dt, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
      c.setUint16(12, tm, true); c.setUint16(14, dt, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + data.length; cd += 46 + name.length;
    });
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cd, true); e.setUint32(16, off, true);
    var all = parts.concat(central, [new Uint8Array(e.buffer)]), len = all.reduce(function (s, a) { return s + a.length; }, 0), out = new Uint8Array(len), p = 0;
    all.forEach(function (a) { out.set(a, p); p += a.length; });
    return out;
  }
  function xe(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''); }
  function colL(c) { var s = ''; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function A1(r, c) { return colL(c) + (r + 1); }
  function absR(sheet, r1, c1, r2, c2) { var q = "'" + sheet.replace(/'/g, "''") + "'!"; return q + '$' + colL(c1) + '$' + (r1 + 1) + (r2 == null ? '' : ':$' + colL(c2) + '$' + (r2 + 1)); }
  function norm(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim(); }
  function round(v, d) { var m = Math.pow(10, d || 0); return Math.round(v * m) / m; }
  function excelDate(iso) { var d = new Date(iso); return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) / 86400000) + 25569; }
  function dayOnly(iso) { return Math.floor(excelDate(iso)); }
  function timeOnly(iso) { var x = excelDate(iso); return x - Math.floor(x); }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function fmtDate(iso) { var d = new Date(iso); return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear(); }

  /* ===================== identidade visual ===================== */
  var C = { verde: '13301A', verde2: '173A1E', bronze: 'A88552', salvia: '8FA37F', creme: 'FBF8F1', creme2: 'EFE8D8', grupo: 'E2DCCB', linha: 'E2D6BC', texto: '3B3F3A', cinza: '899287', branco: 'FFFFFF', ouro: 'B08854' };
  var FONT = 'Century Gothic', SERIF = 'Garamond';
  var CONTATO = '(12) 99182-1729  ·  lucasaosf@gmail.com  ·  @lucas_aosf  ·  CRBio 127233/01-D';
  var SERIES_CORES = [C.verde, C.bronze, C.salvia, 'C9B48A', '5E6B57', 'D6C7A1', '2F5A3A', '7A5A2E'];

  /* ===================== estilos ===================== */
  function Styles() {
    this.fonts = ['<font><sz val="9"/><color rgb="FF' + C.texto + '"/><name val="' + FONT + '"/></font>'];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = []; this.nfId = {}; this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']; this.cache = {};
  }
  Styles.prototype.idx = function (arr, xml) { var i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
  Styles.prototype.get = function (sp) {
    if (!sp) return 0;
    var key = JSON.stringify(sp); if (this.cache[key] != null) return this.cache[key];
    var f = sp.f || {};
    var font = '<font>' + (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.u ? '<u/>' : '') + '<sz val="' + (f.sz || 9) + '"/><color rgb="FF' + (f.c || C.texto) + '"/><name val="' + (f.name || FONT) + '"/></font>';
    var fontId = this.idx(this.fonts, font);
    var fillId = sp.fill ? this.idx(this.fills, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + sp.fill + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
    var borderId = 0;
    if (sp.bd) {
      var bc = sp.bc || C.linha, side = function (n) { return '<' + n + ' style="' + sp.bd + '"><color rgb="FF' + bc + '"/></' + n + '>'; };
      var b = '<border>' + side('left') + side('right') + side('top') + side('bottom') + '<diagonal/></border>';
      if (sp.bdBottomOnly) b = '<border><left/><right/><top/>' + side('bottom') + '<diagonal/></border>';
      borderId = this.idx(this.borders, b);
    }
    var nf = 0;
    if (sp.nf) {
      var builtin = { '0': 1, '0.00': 2, '#,##0': 3, '0%': 9, '0.00%': 10 };
      if (builtin[sp.nf]) nf = builtin[sp.nf];
      else { if (!this.nfId[sp.nf]) { this.nfId[sp.nf] = 164 + this.numFmts.length; this.numFmts.push('<numFmt numFmtId="' + this.nfId[sp.nf] + '" formatCode="' + xe(sp.nf) + '"/>'); } nf = this.nfId[sp.nf]; }
    }
    var al = '';
    if (sp.h || sp.v || sp.wrap || sp.rot || sp.indent) al = '<alignment' + (sp.h ? ' horizontal="' + sp.h + '"' : '') + ' vertical="' + (sp.v || 'center') + '"' + (sp.wrap ? ' wrapText="1"' : '') + (sp.rot ? ' textRotation="' + sp.rot + '"' : '') + (sp.indent ? ' indent="' + sp.indent + '"' : '') + '/>';
    var xf = '<xf numFmtId="' + nf + '" fontId="' + fontId + '" fillId="' + fillId + '" borderId="' + borderId + '" xfId="0"' + (nf ? ' applyNumberFormat="1"' : '') + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '') + (borderId ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>');
    var id = this.idx(this.xfs, xf); this.cache[key] = id; return id;
  };
  Styles.prototype.xml = function () {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (this.numFmts.length ? '<numFmts count="' + this.numFmts.length + '">' + this.numFmts.join('') + '</numFmts>' : '') +
      '<fonts count="' + this.fonts.length + '">' + this.fonts.join('') + '</fonts><fills count="' + this.fills.length + '">' + this.fills.join('') + '</fills>' +
      '<borders count="' + this.borders.length + '">' + this.borders.join('') + '</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join('') + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  };
  // estilos nomeados da marca
  var S = {
    titulo: { f: { b: true, sz: 12, c: C.verde, name: SERIF } },
    subtitulo: { f: { i: true, sz: 9, c: C.cinza } },
    legenda: { f: { sz: 8, c: C.texto }, wrap: true, v: 'top', h: 'left', fill: C.creme2 },
    secao: { f: { b: true, sz: 10, c: C.verde }, bdBottomOnly: true, bd: 'thin', bc: C.bronze },
    th: { f: { b: true, sz: 8.5, c: C.branco }, fill: C.verde, h: 'center', v: 'center', wrap: true, bd: 'thin', bc: C.linha },
    th2: { f: { b: true, sz: 8.5, c: C.verde }, fill: C.grupo, h: 'center', v: 'center', wrap: true, bd: 'thin', bc: C.linha },
    td: { fill: C.creme, h: 'center', bd: 'thin', bc: C.linha },
    tdl: { fill: C.creme, h: 'left', bd: 'thin', bc: C.linha, indent: 1 },
    sp: { f: { i: true }, fill: C.creme, h: 'left', bd: 'thin', bc: C.linha, indent: 1 },
    grp: { f: { b: true, c: C.verde }, fill: C.grupo, h: 'left', bd: 'thin', bc: C.linha, indent: 1 },
    tot: { f: { b: true, c: C.verde }, fill: C.creme2, h: 'center', bd: 'thin', bc: C.linha },
    totl: { f: { b: true, c: C.verde }, fill: C.creme2, h: 'left', bd: 'thin', bc: C.linha, indent: 1 },
    kpiNum: { f: { b: true, sz: 16, c: C.verde, name: SERIF }, h: 'center', fill: C.creme, bd: 'thin', bc: C.linha },
    kpiLab: { f: { sz: 8, c: C.cinza }, h: 'center', fill: C.creme, bd: 'thin', bc: C.linha, wrap: true },
    nota: { f: { i: true, sz: 8, c: C.cinza }, h: 'left', wrap: true, v: 'top' }
  };
  function st(name, extra) { var o = JSON.parse(JSON.stringify(S[name])); if (extra) for (var k in extra) o[k] = extra[k]; return o; }

  /* ===================== planilha (aba) ===================== */
  function Sheet(book, name, opt) {
    this.book = book; this.name = name.slice(0, 31); this.rows = {}; this.cols = {}; this.merges = []; this.charts = []; this.heights = {};
    this.opt = opt || {}; this.maxR = 0; this.maxC = 0;
  }
  Sheet.prototype.set = function (r, c, v, style, formula) {
    if (!this.rows[r]) this.rows[r] = {};
    this.rows[r][c] = { v: v, s: style ? this.book.styles.get(style) : 0, f: formula };
    if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c;
    return this;
  };
  Sheet.prototype.fill = function (r1, c1, r2, c2, style) { for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) if (!this.rows[r] || !this.rows[r][c]) this.set(r, c, null, style); };
  Sheet.prototype.merge = function (r1, c1, r2, c2, style) { if (style) this.fill(r1, c1, r2, c2, style); if (r1 !== r2 || c1 !== c2) this.merges.push(A1(r1, c1) + ':' + A1(r2, c2)); return this; };
  Sheet.prototype.width = function (c, w) { this.cols[c] = w; return this; };
  Sheet.prototype.height = function (r, h) { this.heights[r] = h; return this; };
  Sheet.prototype.chart = function (spec) { this.charts.push(spec); return this; };
  Sheet.prototype.pageBreak = function (r) { (this.breaks = this.breaks || []).push(r); return this; };
  Sheet.prototype.title = function (r, text, sub, span) {
    this.set(r, 0, text, S.titulo); if (span) this.merge(r, 0, r, span - 1);
    if (sub) { this.set(r + 1, 0, sub, S.subtitulo); if (span) this.merge(r + 1, 0, r + 1, span - 1); }
    this.height(r, 20);
    return r + (sub ? 3 : 2);
  };
  Sheet.prototype.xml = function (sheetIndex, rels) {
    var self = this, o = [];
    o.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
    o.push('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    o.push('<dimension ref="A1:' + A1(Math.max(this.maxR, 0), Math.max(this.maxC, 0)) + '"/>');
    o.push('<sheetViews><sheetView workbookViewId="0" showGridLines="0"' + (sheetIndex === 0 ? ' tabSelected="1"' : '') + '>');
    if (this.opt.freeze) { var fr = this.opt.freeze; o.push('<pane' + (fr.c ? ' xSplit="' + fr.c + '"' : '') + (fr.r ? ' ySplit="' + fr.r + '"' : '') + ' topLeftCell="' + A1(fr.r || 0, fr.c || 0) + '" activePane="' + (fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight') + '" state="frozen"/>'); }
    o.push('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/>');
    var cks = Object.keys(this.cols).map(Number).sort(function (a, b) { return a - b; });
    if (cks.length) { o.push('<cols>'); cks.forEach(function (c) { o.push('<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + self.cols[c] + '" customWidth="1"/>'); }); o.push('</cols>'); }
    o.push('<sheetData>');
    Object.keys(this.rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      var row = self.rows[r], h = self.heights[r];
      o.push('<row r="' + (r + 1) + '"' + (h ? ' ht="' + h + '" customHeight="1"' : '') + '>');
      Object.keys(row).map(Number).sort(function (a, b) { return a - b; }).forEach(function (c) {
        var cell = row[c], ref = A1(r, c), s = cell.s ? ' s="' + cell.s + '"' : '', v = cell.v;
        if (cell.f) { o.push('<c r="' + ref + '"' + s + '><f>' + xe(cell.f) + '</f>' + (typeof v === 'number' && isFinite(v) ? '<v>' + v + '</v>' : '') + '</c>'); }
        else if (v == null || v === '') o.push('<c r="' + ref + '"' + s + '/>');
        else if (typeof v === 'number') o.push(isFinite(v) ? '<c r="' + ref + '"' + s + '><v>' + v + '</v></c>' : '<c r="' + ref + '"' + s + '/>');
        else o.push('<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + xe(v) + '</t></is></c>');
      });
      o.push('</row>');
    });
    o.push('</sheetData>');
    if (this.merges.length) o.push('<mergeCells count="' + this.merges.length + '">' + this.merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>');
    o.push('<pageMargins left="0.51" right="0.51" top="0.98" bottom="0.79" header="0.3" footer="0.3"/>');
    o.push('<pageSetup paperSize="9" orientation="' + (this.opt.landscape ? 'landscape' : 'portrait') + '" fitToWidth="' + (this.opt.fitW || 1) + '" fitToHeight="0"/>');
    var tab = xe(this.name).replace(/&amp;/g, '&amp;&amp;');
    o.push('<headerFooter><oddHeader>' + xe('&L&"' + SERIF + ',Regular"&13&K' + C.verde + 'Lucas Aøsf\n&"' + FONT + ',Regular"&6&K' + C.cinza + 'BIÓLOGO · PESQUISADOR · GESTOR AMBIENTAL&C&"' + FONT + ',Bold"&8&K' + C.verde + '&A\n&"' + FONT + ',Italic"&7&K' + C.cinza + 'Pesquisa e Soluções Ambientais&R&G') +
      '</oddHeader><oddFooter>' + xe('&L&"' + FONT + ',Regular"&7&K' + C.cinza + CONTATO + '&R&"' + SERIF + ',Regular"&11&K' + C.verde + '&P&"' + FONT + ',Regular"&7&K' + C.cinza + ' / &N') + '</oddFooter></headerFooter>');
    if (this.breaks && this.breaks.length) o.push('<rowBreaks count="' + this.breaks.length + '" manualBreakCount="' + this.breaks.length + '">' + this.breaks.map(function (b) { return '<brk id="' + b + '" max="16383" man="1"/>'; }).join('') + '</rowBreaks>');
    if (this.charts.length) o.push('<drawing r:id="' + rels.drawing + '"/>');
    if (rels.vml) o.push('<legacyDrawingHF r:id="' + rels.vml + '"/>');
    o.push('</worksheet>');
    return o.join('');
  };

  /* ===================== gráficos ===================== */
  function txPr(sz, color, bold) { return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + (sz || 900) + '" b="' + (bold ? 1 : 0) + '"><a:solidFill><a:srgbClr val="' + (color || C.texto) + '"/></a:solidFill><a:latin typeface="' + FONT + '"/></a:defRPr></a:pPr><a:endParaRPr lang="pt-BR"/></a:p></c:txPr>'; }
  function richTitle(text, sz, color) { return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + sz + '" b="0"><a:solidFill><a:srgbClr val="' + (color || C.verde) + '"/></a:solidFill><a:latin typeface="' + FONT + '"/></a:defRPr></a:pPr><a:r><a:rPr lang="pt-BR" sz="' + sz + '" b="0"><a:solidFill><a:srgbClr val="' + (color || C.verde) + '"/></a:solidFill><a:latin typeface="' + FONT + '"/></a:rPr><a:t>' + xe(text) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'; }
  function strCache(vals) { return '<c:strCache><c:ptCount val="' + vals.length + '"/>' + vals.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + xe(v) + '</c:v></c:pt>'; }).join('') + '</c:strCache>'; }
  function numCache(vals) { return '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + vals.length + '"/>' + vals.map(function (v, i) { return (v == null || !isFinite(v)) ? '' : '<c:pt idx="' + i + '"><c:v>' + v + '</c:v></c:pt>'; }).join('') + '</c:numCache>'; }
  function serXml(s, i, kind, catRef, catVals) {
    var color = s.color || SERIES_CORES[i % SERIES_CORES.length], o = '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/>';
    o += '<c:tx>' + (s.nameRef ? '<c:strRef><c:f>' + xe(s.nameRef) + '</c:f>' + strCache([s.name]) + '</c:strRef>' : '<c:v>' + xe(s.name) + '</c:v>') + '</c:tx>';
    if (kind === 'line') {
      o += '<c:spPr><a:ln w="' + (s.width || 25400) + '" cap="rnd"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' + (s.dash ? '<a:prstDash val="' + s.dash + '"/>' : '') + '<a:round/></a:ln></c:spPr>';
      o += s.marker === false ? '<c:marker><c:symbol val="none"/></c:marker>' : '<c:marker><c:symbol val="' + (s.marker || 'circle') + '"/><c:size val="' + (s.msize || 6) + '"/><c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></a:ln></c:spPr></c:marker>';
    } else if (kind === 'pie') {
      o += '<c:spPr><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>';
      (s.vals || []).forEach(function (_, k) { o += '<c:dPt><c:idx val="' + k + '"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="' + SERIES_CORES[k % SERIES_CORES.length] + '"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>'; });
    } else {
      o += '<c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>';
    }
    if (s.labels) {
      o += '<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' + txPr(800, C.texto, true) + (kind === 'pie' ? '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:separator>: </c:separator><c:showLeaderLines val="1"/>' : '<c:dLblPos val="' + (kind === 'line' ? 't' : 'outEnd') + '"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>') + '</c:dLbls>';
    }
    if (s.err) {
      o += '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="0"/>' +
        '<c:plus><c:numRef><c:f>' + xe(s.err.ref) + '</c:f>' + numCache(s.err.vals) + '</c:numRef></c:plus><c:minus><c:numRef><c:f>' + xe(s.err.ref) + '</c:f>' + numCache(s.err.vals) + '</c:numRef></c:minus>' +
        '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></a:ln></c:spPr></c:errBars>';
    }
    o += '<c:cat>' + (catRef.num ? '<c:numRef><c:f>' + xe(catRef.ref) + '</c:f>' + numCache(catVals) + '</c:numRef>' : '<c:strRef><c:f>' + xe(catRef.ref) + '</c:f>' + strCache(catVals) + '</c:strRef>') + '</c:cat>';
    o += '<c:val><c:numRef><c:f>' + xe(s.ref) + '</c:f>' + numCache(s.vals) + '</c:numRef></c:val>';
    if (kind === 'line') o += '<c:smooth val="0"/>';
    return o + '</c:ser>';
  }
  function axes(spec, catId, valId) {
    var grid = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E6E1D3"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
    var axLn = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFB8A6"/></a:solidFill></a:ln></c:spPr>';
    return '<c:catAx><c:axId val="' + catId + '"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' + (spec.xTitle ? richTitle(spec.xTitle, 900, C.texto) : '') +
      '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' + axLn + txPr(800, C.texto) + '<c:crossAx val="' + valId + '"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>' + (spec.skip ? '<c:tickLblSkip val="' + spec.skip + '"/>' : '') + '<c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="' + valId + '"/><c:scaling><c:orientation val="minMax"/>' + '<c:min val="' + (spec.yMin != null ? spec.yMin : 0) + '"/>' + '</c:scaling><c:delete val="0"/><c:axPos val="l"/>' + grid + (spec.yTitle ? richTitle(spec.yTitle, 900, C.texto) : '') +
      '<c:numFmt formatCode="' + (spec.yFmt || 'General') + '" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>' + txPr(800, C.texto) + '<c:crossAx val="' + catId + '"/><c:crosses val="autoZero"/><c:crossBetween val="' + (spec.type === 'line' ? 'midCat' : 'between') + '"/>' + (spec.yInt ? '<c:majorUnit val="' + spec.yInt + '"/>' : '') + '</c:valAx>';
  }
  function chartXml(spec) {
    var o = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="pt-BR"/><c:roundedCorners val="0"/><c:chart>';
    o += richTitle(spec.title, 1300) + '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>';
    var cat = { ref: spec.cats.ref, num: spec.cats.num }, cv = spec.cats.vals;
    if (spec.type === 'pie') {
      o += '<c:pieChart><c:varyColors val="1"/>' + spec.series.map(function (s, i) { return serXml(s, i, 'pie', cat, cv); }).join('') + '<c:firstSliceAng val="0"/></c:pieChart>';
    } else {
      var bars = spec.series.filter(function (s) { return (s.kind || (spec.type === 'line' ? 'line' : 'bar')) === 'bar'; });
      var lines = spec.series.filter(function (s) { return (s.kind || (spec.type === 'line' ? 'line' : 'bar')) === 'line'; });
      var i = 0;
      if (bars.length) o += '<c:barChart><c:barDir val="' + (spec.horizontal ? 'bar' : 'col') + '"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + bars.map(function (s) { return serXml(s, i++, 'bar', cat, cv); }).join('') + '<c:gapWidth val="' + (spec.gap || 80) + '"/><c:overlap val="' + (spec.overlap || 0) + '"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>';
      if (lines.length) o += '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + lines.map(function (s) { return serXml(s, i++, 'line', cat, cv); }).join('') + '<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>';
      o += axes(spec, 5001, 5002);
      if (spec.horizontal) o = o.replace('<c:axPos val="b"/>', '<c:axPos val="l"/>').replace(/<c:valAx>(.*?)<c:axPos val="l"\/>/, '<c:valAx>$1<c:axPos val="b"/>');
    }
    o += '<c:spPr><a:noFill/></c:spPr></c:plotArea>';
    if (spec.legend !== false) o += '<c:legend><c:legendPos val="' + (spec.legend || 'b') + '"/><c:overlay val="0"/>' + txPr(800, C.texto) + '</c:legend>';
    o += '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D2C1"/></a:solidFill></a:ln></c:spPr>' + txPr(900, C.texto) + '</c:chartSpace>';
    return o;
  }
  function drawingXml(charts, firstId) {
    var o = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
    charts.forEach(function (ch, i) {
      var a = ch.at;
      o += '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + a.c1 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + a.r1 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + a.c2 + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + a.r2 + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Gráfico ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + (i + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
    });
    return o + '</xdr:wsDr>';
  }
  var VML_HF = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="RH" o:spid="_x0000_s1025" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:30.9pt;height:32pt;z-index:1"><v:imagedata o:relid="rId1" o:title="Logo"/><o:lock v:ext="edit" rotation="t"/></v:shape></xml>';

  /* ===================== pasta de trabalho ===================== */
  function Book(meta) { this.sheets = []; this.styles = new Styles(); this.meta = meta || {}; this.logo = null; }
  Book.prototype.sheet = function (name, opt) { var s = new Sheet(this, name, opt); this.sheets.push(s); return s; };
  Book.prototype.build = function () {
    var self = this, files = [], ct = [], wbRels = [], chartN = 0, drawN = 0;
    var sheetsXml = this.sheets.map(function (s, i) {
      var rels = [], relIds = {};
      if (s.charts.length) {
        drawN++; relIds.drawing = 'rId' + (rels.length + 1); rels.push('<Relationship Id="' + relIds.drawing + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + drawN + '.xml"/>');
        var drels = [];
        s.charts.forEach(function (ch, k) {
          chartN++;
          files.push({ name: 'xl/charts/chart' + chartN + '.xml', data: chartXml(ch) });
          ct.push('<Override PartName="/xl/charts/chart' + chartN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
          drels.push('<Relationship Id="rId' + (k + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + chartN + '.xml"/>');
        });
        files.push({ name: 'xl/drawings/drawing' + drawN + '.xml', data: drawingXml(s.charts) });
        files.push({ name: 'xl/drawings/_rels/drawing' + drawN + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + drels.join('') + '</Relationships>' });
        ct.push('<Override PartName="/xl/drawings/drawing' + drawN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
      }
      if (self.logo) {
        relIds.vml = 'rId' + (rels.length + 1);
        rels.push('<Relationship Id="' + relIds.vml + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawingHF' + (i + 1) + '.vml"/>');
        files.push({ name: 'xl/drawings/vmlDrawingHF' + (i + 1) + '.vml', data: VML_HF });
        files.push({ name: 'xl/drawings/_rels/vmlDrawingHF' + (i + 1) + '.vml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.png"/></Relationships>' });
      }
      if (rels.length) files.push({ name: 'xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join('') + '</Relationships>' });
      ct.push('<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      wbRels.push('<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>');
      return s.xml(i, relIds);
    });
    sheetsXml.forEach(function (x, i) { files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: x }); });
    if (this.logo) files.push({ name: 'xl/media/logo.png', data: this.logo });
    var n = this.sheets.length;
    wbRels.push('<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');
    var names = this.sheets.map(function (s, i) { return '<sheet name="' + xe(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('');
    var defNames = this.sheets.map(function (s, i) { return s.opt.printTitles ? '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + xe("'" + s.name.replace(/'/g, "''") + "'!$" + (s.opt.printTitles[0] + 1) + ':$' + (s.opt.printTitles[1] + 1)) + '</definedName>' : ''; }).join('');
    files.push({ name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>' + names + '</sheets>' + (defNames ? '<definedNames>' + defNames + '</definedNames>' : '') + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>' });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join('') + '</Relationships>' });
    files.push({ name: 'xl/styles.xml', data: this.styles.xml() });
    var m = this.meta, now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    files.push({ name: 'docProps/core.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xe(m.title || 'Lista de Campo') + '</dc:title><dc:creator>' + xe(m.author || 'Lista de Campo · Lucas Aøsf') + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created></cp:coreProperties>' });
    files.push({ name: 'docProps/app.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Lista de Campo · Lucas Aøsf</Application></Properties>' });
    files.unshift({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>' });
    files.unshift({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + ct.join('') + '</Types>' });
    return zip(files);
  };

  /* ===================== análises ===================== */
  function lnfact(n) { var s = 0; for (var i = 2; i <= n; i++) s += Math.log(i); return s; }
  function lnC(n, k) { if (k < 0 || k > n) return -Infinity; return lnfact(n) - lnfact(k) - lnfact(n - k); }
  function shannon(counts) { var N = counts.reduce(function (a, b) { return a + b; }, 0); if (!N) return 0; return -counts.reduce(function (h, n) { if (!n) return h; var p = n / N; return h + p * Math.log(p); }, 0); }
  function simpson(counts) { var N = counts.reduce(function (a, b) { return a + b; }, 0); if (N < 2) return null; return 1 - counts.reduce(function (s, n) { return s + n * (n - 1); }, 0) / (N * (N - 1)); }
  function chao1(counts) { var S = counts.filter(Boolean).length, F1 = counts.filter(function (n) { return n === 1; }).length, F2 = counts.filter(function (n) { return n === 2; }).length; return S + F1 * (F1 - 1) / (2 * (F2 + 1)); }
  function incStats(mat) { // mat: amostras × espécies (presença)
    var m = mat.length, q = {}, S = 0;
    if (!m) return { m: 0, S: 0, Q1: 0, Q2: 0, inc: [] };
    var nsp = mat[0].length, inc = [];
    for (var j = 0; j < nsp; j++) { var k = 0; for (var i = 0; i < m; i++) if (mat[i][j] > 0) k++; inc.push(k); if (k) S++; }
    var Q1 = inc.filter(function (k) { return k === 1; }).length, Q2 = inc.filter(function (k) { return k === 2; }).length;
    return { m: m, S: S, Q1: Q1, Q2: Q2, inc: inc };
  }
  function chao2(st) { if (!st.m) return 0; var m = st.m; return st.S + ((m - 1) / m) * st.Q1 * (st.Q1 - 1) / (2 * (st.Q2 + 1)); }
  function jack1(st) { if (!st.m) return 0; return st.S + st.Q1 * (st.m - 1) / st.m; }
  function jack2(st) { var m = st.m; if (m < 2) return st.S; return st.S + st.Q1 * (2 * m - 3) / m - st.Q2 * Math.pow(m - 2, 2) / (m * (m - 1)); }
  function maoTau(st) { // rarefação por amostras (Colwell et al. 2004) com IC 95%
    var m = st.m, out = [], Sest = Math.max(chao2(st), st.S);
    var sk = {}; st.inc.forEach(function (k) { if (k) sk[k] = (sk[k] || 0) + 1; });
    for (var h = 1; h <= m; h++) {
      var Sh = st.S, v = 0;
      Object.keys(sk).forEach(function (k) {
        k = +k; var a = (k <= m - h) ? Math.exp(lnC(m - k, h) - lnC(m, h)) : 0;
        Sh -= a * sk[k]; v += Math.pow(1 - a, 2) * sk[k];
      });
      v -= Sh * Sh / Sest; var sd = Math.sqrt(Math.max(v, 0));
      out.push({ h: h, S: Sh, sd: sd, lo: Math.max(Sh - 1.96 * sd, 0), hi: Sh + 1.96 * sd });
    }
    return out;
  }
  function rng(seed) { var s = seed >>> 0 || 1; return function () { s ^= s << 13; s >>>= 0; s ^= s >>> 17; s ^= s << 5; s >>>= 0; return s / 4294967296; }; }
  function jackCurve(mat, runs) { // Jackknife 1 médio por nº de amostras (aleatorizações)
    var m = mat.length, sum = [], sq = [], r = rng(12345);
    for (var h = 0; h < m; h++) { sum.push(0); sq.push(0); }
    for (var t = 0; t < runs; t++) {
      var ord = mat.map(function (_, i) { return i; });
      for (var i = m - 1; i > 0; i--) { var j = Math.floor(r() * (i + 1)), x = ord[i]; ord[i] = ord[j]; ord[j] = x; }
      for (var hh = 1; hh <= m; hh++) { var v = jack1(incStats(ord.slice(0, hh).map(function (k) { return mat[k]; }))); sum[hh - 1] += v; sq[hh - 1] += v * v; }
    }
    return sum.map(function (s, i) { var mean = s / runs; return { h: i + 1, mean: mean, sd: Math.sqrt(Math.max(sq[i] / runs - mean * mean, 0)) }; });
  }
  function brayCurtisSim(a, b) { var num = 0, den = 0; for (var i = 0; i < a.length; i++) { num += Math.abs(a[i] - b[i]); den += a[i] + b[i]; } return den ? 1 - num / den : null; }
  function jaccard(a, b) { var A = 0, B = 0, Cc = 0; for (var i = 0; i < a.length; i++) { var x = a[i] > 0, y = b[i] > 0; if (x && y) A++; else if (x) B++; else if (y) Cc++; } return (A + B + Cc) ? A / (A + B + Cc) : null; }
  function utm(lat, lon) {
    if (lat == null || lon == null) return null;
    var zone = Math.floor((lon + 180) / 6) + 1, a = 6378137, f = 1 / 298.257223563, k0 = 0.9996, e2 = f * (2 - f), ep2 = e2 / (1 - e2);
    var lam0 = ((zone - 1) * 6 - 180 + 3) * Math.PI / 180, phi = lat * Math.PI / 180, lam = lon * Math.PI / 180;
    var N = a / Math.sqrt(1 - e2 * Math.sin(phi) * Math.sin(phi)), T = Math.tan(phi) * Math.tan(phi), Cc = ep2 * Math.cos(phi) * Math.cos(phi), A = Math.cos(phi) * (lam - lam0);
    var M = a * ((1 - e2 / 4 - 3 * e2 * e2 / 64 - 5 * e2 * e2 * e2 / 256) * phi - (3 * e2 / 8 + 3 * e2 * e2 / 32 + 45 * e2 * e2 * e2 / 1024) * Math.sin(2 * phi) + (15 * e2 * e2 / 256 + 45 * e2 * e2 * e2 / 1024) * Math.sin(4 * phi) - (35 * e2 * e2 * e2 / 3072) * Math.sin(6 * phi));
    var x = k0 * N * (A + (1 - T + Cc) * Math.pow(A, 3) / 6 + (5 - 18 * T + T * T + 72 * Cc - 58 * ep2) * Math.pow(A, 5) / 120) + 500000;
    var y = k0 * (M + N * Math.tan(phi) * (A * A / 2 + (5 - T + 9 * Cc + 4 * Cc * Cc) * Math.pow(A, 4) / 24 + (61 - 58 * T + T * T + 600 * Cc - 330 * ep2) * Math.pow(A, 6) / 720));
    if (lat < 0) y += 10000000;
    var bands = 'CDEFGHJKLMNPQRSTUVWX', band = bands.charAt(Math.max(0, Math.min(19, Math.floor((lat + 80) / 8))));
    return { zone: zone, band: band, e: Math.round(x), n: Math.round(y) };
  }

  /* ===================== preparação dos dados ===================== */
  var CLASSE_ORDEM = { 'Anfíbio': 1, 'Réptil': 2, 'Ave': 3, 'Mamífero': 4 };
  var CLASSE_NOME = { 'Anfíbio': 'Amphibia', 'Réptil': 'Reptilia', 'Ave': 'Aves', 'Mamífero': 'Mammalia' };
  function preparar(input) {
    var listas = (input.listas || []).slice().sort(function (a, b) { return a.inicio < b.inicio ? -1 : 1; });
    var db = input.especie || function () { return null; };
    var spMap = {}, spList = [];
    var registros = [];
    listas.forEach(function (L, li) {
      L._i = li;
      L.ua = (L.ua || '').trim() || 'Sem unidade';
      L.campanha = (L.campanha || '').trim() || fmtDate(L.inicio);
      L.metodo = (L.metodo || '').trim() || 'Não informado';
      L.registros.forEach(function (r) {
        var key = r.sci || r.nome; if (!key) return;
        var e = spMap[key];
        if (!e) {
          var d = r.sci ? db(r.sci) : null;
          e = spMap[key] = { key: key, sci: r.sci || '', nome: r.sci || r.nome, livre: !r.sci, d: d || {}, autor: d ? d.a : '', N: 0, v: 0, a: 0, o: 0 };
          spList.push(e);
        }
        var tot = (r.v || 0) + (r.a || 0) + (r.outro || 0);
        if (!tot) return;
        e.N += tot; e.v += r.v || 0; e.a += r.a || 0; e.o += r.outro || 0;
        registros.push({ L: L, e: e, r: r, tot: tot });
      });
    });
    spList = spList.filter(function (e) { return e.N > 0; });
    spList.sort(function (x, y) {
      var cx = CLASSE_ORDEM[x.d.c] || 9, cy = CLASSE_ORDEM[y.d.c] || 9; if (cx !== cy) return cx - cy;
      var o = (x.d.o || '~').localeCompare(y.d.o || '~'); if (o) return o;
      var f = (x.d.f || '~').localeCompare(y.d.f || '~'); if (f) return f;
      return x.nome.localeCompare(y.nome);
    });
    spList.forEach(function (e, i) { e.i = i; });
    var uas = [], uaIdx = {}, camps = [], campIdx = {}, mets = [], metIdx = {};
    listas.forEach(function (L) {
      if (uaIdx[L.ua] == null) { uaIdx[L.ua] = uas.length; uas.push({ nome: L.ua, listas: [], lat: null, lon: null }); }
      var u = uas[uaIdx[L.ua]]; u.listas.push(L); if (u.lat == null && L.lat != null) { u.lat = L.lat; u.lon = L.lon; }
      if (campIdx[L.campanha] == null) { campIdx[L.campanha] = camps.length; camps.push({ nome: L.campanha, listas: [], inicio: L.inicio, fim: L.fim || L.inicio }); }
      var c = camps[campIdx[L.campanha]]; c.listas.push(L); if ((L.fim || L.inicio) > c.fim) c.fim = L.fim || L.inicio;
      if (metIdx[L.metodo] == null) { metIdx[L.metodo] = mets.length; mets.push({ nome: L.metodo }); }
    });
    uas.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt', { numeric: true }); }); uas.forEach(function (u, i) { uaIdx[u.nome] = i; });
    var nS = spList.length;
    function zeros() { var z = []; for (var i = 0; i < nS; i++) z.push(0); return z; }
    var byUA = uas.map(function () { return { v: zeros(), a: zeros(), t: zeros() }; });
    var byCamp = camps.map(zeros), byList = listas.map(zeros), byMet = mets.map(zeros);
    registros.forEach(function (x) {
      var i = x.e.i, u = byUA[uaIdx[x.L.ua]];
      u.v[i] += x.r.v || 0; u.a[i] += x.r.a || 0; u.t[i] += x.tot;
      byCamp[campIdx[x.L.campanha]][i] += x.tot; byList[x.L._i][i] += x.tot; byMet[metIdx[x.L.metodo]][i] += x.tot;
    });
    var Ntot = spList.reduce(function (s, e) { return s + e.N; }, 0);
    return { listas: listas, sp: spList, registros: registros, uas: uas, camps: camps, mets: mets, byUA: byUA, byCamp: byCamp, byList: byList, byMet: byMet, N: Ntot, input: input };
  }
  function estimativa(pct) { return pct >= 5 ? 'M' : pct >= 1 ? 'A' : 'P'; }
  function constancia(fo) { return fo >= 50 ? 'Constante' : fo >= 25 ? 'Acessória' : 'Acidental'; }
  function tipoReg(e) { var t = []; if (e.v) t.push('V'); if (e.a) t.push('A'); if (!t.length && e.o) t.push('—'); return t.join(', '); }
  var IUCN_NOME = { LC: 'Pouco preocupante', NT: 'Quase ameaçada', VU: 'Vulnerável', EN: 'Em perigo', CR: 'Criticamente em perigo', EW: 'Extinta na natureza', EX: 'Extinta', DD: 'Dados insuficientes', 'N/A': 'Não avaliada' };
  var AMEACA = { VU: 1, EN: 1, CR: 1, EW: 1, EX: 1 };

  /* ===================== montagem da planilha ===================== */
  function gerar(input) {
    var D = preparar(input), B = new Book({ title: input.titulo, author: 'Lista de Campo · Lucas Aøsf' });
    if (input.logo) B.logo = input.logo;
    var nS = D.sp.length, nU = D.uas.length, nC = D.camps.length, nL = D.listas.length;
    var titulo = input.titulo || 'Levantamento de fauna';
    var dIni = D.listas.length ? fmtDate(D.listas[0].inicio) : '', dFim = D.listas.length ? fmtDate(D.listas[nL - 1].fim || D.listas[nL - 1].inicio) : '';
    var periodo = dIni === dFim ? dIni : dIni + ' a ' + dFim;
    var grupos = {}; D.sp.forEach(function (e) { grupos[e.d.c || 'Outros'] = 1; });
    var gruposTxt = Object.keys(grupos).map(function (g) { return g === 'Outros' ? 'sem identificação no banco' : g.toLowerCase(); }).join(', ');
    var horas = D.listas.reduce(function (s, L) { return s + (L.durMin || 0); }, 0) / 60, km = D.listas.reduce(function (s, L) { return s + (L.distKm || 0); }, 0);
    var incAll = incStats(D.byList), totCounts = D.sp.map(function (e) { return e.N; });
    var H = shannon(totCounts), J = nS > 1 ? H / Math.log(nS) : null, Simp = simpson(totCounts);

    /* ---------- 1. Resumo ---------- */
    var R = B.sheet('Resumo', { landscape: false });
    [34, 15, 15, 15, 15, 15].forEach(function (w, i) { R.width(i, w); });
    var r = R.title(0, titulo, 'Planilha gerada pelo app Lista de Campo · ' + periodo, 6);
    R.set(r, 0, 'Dados do levantamento', S.secao); R.fill(r, 1, r, 5, S.secao); r += 1;
    var info = [['Projeto / área', input.projeto || (D.listas[0] && D.listas[0].projeto) || '—'], ['Campanhas', D.camps.map(function (c) { return c.nome; }).join(', ')], ['Período', periodo], ['Unidades amostrais', D.uas.map(function (u) { return u.nome; }).join(', ')], ['Métodos', D.mets.map(function (m) { return m.nome; }).join(', ')], ['Grupos registrados', gruposTxt || '—'], ['Observadores (máx. por lista)', Math.max.apply(null, D.listas.map(function (L) { return L.observadores || 1; }).concat([1]))]];
    info.forEach(function (x) { R.set(r, 0, x[0], S.totl); R.set(r, 1, x[1], st('tdl', { wrap: true })); R.merge(r, 1, r, 5, st('tdl')); r++; });
    r++;
    R.set(r, 0, 'Resultados', S.secao); R.fill(r, 1, r, 5, S.secao); r += 1;
    var kpis = [[nS, 'Riqueza (espécies)'], [D.N, 'Abundância (indivíduos)'], [nL, 'Listas / amostras'], [round(horas, 1), 'Esforço (horas)'], [round(km, 2), 'Distância (km)'], [nU, 'Unidades amostrais']];
    kpis.forEach(function (k, i) { R.set(r, i, k[0], st('kpiNum', { nf: typeof k[0] === 'number' && k[0] % 1 ? '0.0' : '0' })); R.set(r + 1, i, k[1], S.kpiLab); });
    R.height(r, 28); R.height(r + 1, 24); r += 3;
    var idx = [['Diversidade de Shannon (H\')', round(H, 3)], ['Equitabilidade de Pielou (J\')', J == null ? '—' : round(J, 3)], ['Diversidade de Simpson (1 − D)', Simp == null ? '—' : round(Simp, 3)], ['Riqueza estimada — Chao 1', round(chao1(totCounts), 1)], ['Riqueza estimada — Chao 2', nL > 1 ? round(chao2(incAll), 1) : '—'], ['Riqueza estimada — Jackknife 1', nL > 1 ? round(jack1(incAll), 1) : '—'], ['Riqueza estimada — Jackknife 2', nL > 2 ? round(jack2(incAll), 1) : '—'], ['Completude da amostragem (S obs / Jackknife 1)', nL > 1 && jack1(incAll) ? round(100 * nS / jack1(incAll), 1) + '%' : '—']];
    R.set(r, 0, 'Índice', S.th); R.merge(r, 0, r, 3, S.th); R.set(r, 4, 'Valor', S.th); R.merge(r, 4, r, 5, S.th); r++;
    idx.forEach(function (x) { R.set(r, 0, x[0], S.tdl); R.merge(r, 0, r, 3, S.tdl); R.set(r, 4, x[1], S.td); R.merge(r, 4, r, 5, S.td); r++; });
    r++;
    var ameacadas = D.sp.filter(function (e) { return AMEACA[e.d.iucn] || AMEACA[e.d.mma]; }), endem = D.sp.filter(function (e) { return e.d.e === 'Sim'; }), exot = D.sp.filter(function (e) { return e.d.og === 'Exótica' || e.d.og === 'Introduzida'; }), livres = D.sp.filter(function (e) { return e.livre; });
    R.set(r, 0, 'Destaques', S.secao); R.fill(r, 1, r, 5, S.secao); r++;
    [['Espécies ameaçadas (IUCN ou MMA)', ameacadas], ['Espécies endêmicas do Brasil', endem], ['Espécies exóticas ou introduzidas', exot], ['Nomes sem espécie do banco (conferir)', livres]].forEach(function (x) {
      R.set(r, 0, x[0], S.totl); R.set(r, 1, x[1].length, S.td); R.set(r, 2, x[1].map(function (e) { return e.nome; }).join(', ') || '—', st('sp', { wrap: true })); R.merge(r, 2, r, 5, st('sp', { wrap: true }));
      if (x[1].length > 4) R.height(r, Math.min(13.5 * Math.ceil(x[1].length / 4), 120)); r++;
    });
    r++;
    R.set(r, 0, 'Ameaça e endemismo vêm das listas de referência (veja a aba Legenda e fontes). Aves, répteis e anfíbios não têm grau de ameaça nessas listas.', S.nota); R.merge(r, 0, r + 1, 5, S.nota);

    /* ---------- 2. Espécies ---------- */
    var E = B.sheet('Espécies', { landscape: true, freeze: { r: 5, c: 4 }, printTitles: [3, 4], fitW: 2 });
    var uaCols = D.uas.length;
    var headE = ['Classe', 'Ordem', 'Família', 'Espécie', 'Autor', 'Nome popular', 'Tipo de registro'];
    var after = ['N.T', 'Abund. rel. (%)', 'Estimativa', 'F.O. (%)', 'Constância', 'Unidades amostrais', 'Origem', 'Endêmica BR', 'Ocorrência (aves)', 'IUCN', 'MMA', 'CITES', 'Dieta', 'Locomoção', 'UF de ocorrência', 'Biomas'];
    var c0 = headE.length, cAfter = c0 + uaCols * 2;
    r = E.title(0, 'Tabela de espécies registradas', titulo + ' · ' + periodo + ' · V: visual; A: auditivo; N.T: número total de indivíduos; F.O.: frequência de ocorrência nas listas.', Math.min(cAfter + after.length, 14));
    headE.forEach(function (h, i) { E.set(r, i, h, S.th); E.merge(r, i, r + 1, i, S.th); });
    D.uas.forEach(function (u, k) { E.set(r, c0 + k * 2, u.nome, S.th); E.merge(r, c0 + k * 2, r, c0 + k * 2 + 1, S.th); E.set(r + 1, c0 + k * 2, 'V', S.th2); E.set(r + 1, c0 + k * 2 + 1, 'A', S.th2); });
    after.forEach(function (h, i) { E.set(r, cAfter + i, h, S.th); E.merge(r, cAfter + i, r + 1, cAfter + i, S.th); });
    E.height(r, 26); r += 2;
    var rE0 = r;
    D.sp.forEach(function (e) {
      var d = e.d, pct = D.N ? 100 * e.N / D.N : 0, fo = nL ? 100 * D.byList.filter(function (l) { return l[e.i] > 0; }).length / nL : 0;
      var uaPres = D.uas.filter(function (u, k) { return D.byUA[k].t[e.i] > 0; }).map(function (u) { return u.nome; }).join(', ');
      var vals = [CLASSE_NOME[d.c] || '—', d.o || '—', d.f || '—', e.nome, e.autor || '', d.pt ? d.pt.split(',')[0].trim() : (e.livre && e.key !== e.nome ? e.key : ''), tipoReg(e)];
      vals.forEach(function (v, i) { E.set(r, i, v, i === 3 ? S.sp : i === 0 || i === 6 ? S.td : S.tdl); });
      D.uas.forEach(function (u, k) { var v = D.byUA[k].v[e.i], a = D.byUA[k].a[e.i] + (D.byUA[k].t[e.i] - D.byUA[k].v[e.i] - D.byUA[k].a[e.i]); E.set(r, c0 + k * 2, v || null, st('td', { nf: '0' })); E.set(r, c0 + k * 2 + 1, a || null, st('td', { nf: '0' })); });
      var sumF = uaCols ? 'SUM(' + A1(r, c0) + ':' + A1(r, c0 + uaCols * 2 - 1) + ')' : null;
      E.set(r, cAfter, e.N, st('td', { nf: '0' }), sumF);
      E.set(r, cAfter + 1, round(pct, 2), st('td', { nf: '0.00' }), sumF ? A1(r, cAfter) + '/' + A1(rE0 + nS, cAfter) + '*100' : null);
      var rest = [estimativa(pct), round(fo, 1), constancia(fo), uaPres, d.og || (e.livre ? '—' : ''), d.e || '', d.oc || '', d.iucn || '', d.mma || '', d.cites && d.cites !== '-' ? d.cites : '', d.die || '', d.loc || '', d.uf || '', d.bio || ''];
      rest.forEach(function (v, i) { E.set(r, cAfter + 2 + i, v, i === 3 || i >= 12 ? st('tdl', { wrap: false }) : st('td', i === 1 ? { nf: '0.0' } : null)); });
      r++;
    });
    E.set(r, 0, 'Abundância por unidade amostral', S.totl); E.merge(r, 0, r, c0 - 1, S.totl);
    D.uas.forEach(function (u, k) { [0, 1].forEach(function (j) { var col = c0 + k * 2 + j; E.set(r, col, null, st('tot', { nf: '0' }), 'SUM(' + A1(rE0, col) + ':' + A1(r - 1, col) + ')'); }); });
    E.set(r, cAfter, D.N, st('tot', { nf: '0' }), 'SUM(' + A1(rE0, cAfter) + ':' + A1(r - 1, cAfter) + ')'); E.fill(r, cAfter + 1, r, cAfter + after.length - 1, S.tot); r++;
    E.set(r, 0, 'Riqueza por unidade amostral', S.totl); E.merge(r, 0, r, c0 - 1, S.totl);
    D.uas.forEach(function (u, k) { E.set(r, c0 + k * 2, D.byUA[k].t.filter(Boolean).length, st('tot', { nf: '0' })); E.merge(r, c0 + k * 2, r, c0 + k * 2 + 1, S.tot); });
    E.set(r, cAfter, nS, st('tot', { nf: '0' })); E.fill(r, cAfter + 1, r, cAfter + after.length - 1, S.tot); r += 2;
    E.set(r, 0, 'Estimativa de abundância: M – muito abundante (≥ 5% dos indivíduos); A – abundante (1 a 5%); P – pouco abundante (< 1%). Constância (Dajoz): constante ≥ 50% das listas; acessória 25–50%; acidental < 25%. Dieta e locomoção: códigos da SBMz (ver Legenda e fontes).', S.legenda); E.merge(r, 0, r + 2, 12, S.legenda);
    [10, 15, 16, 26, 22, 20, 9].forEach(function (w, i) { E.width(i, w); });
    for (var k2 = 0; k2 < uaCols * 2; k2++) E.width(c0 + k2, 5.5);
    [6, 8, 8, 7, 10, 18, 10, 9, 18, 6, 6, 7, 8, 8, 26, 30].forEach(function (w, i) { E.width(cAfter + i, w); });

    /* ---------- 3. Unidades amostrais (riqueza e abundância) ---------- */
    var U = B.sheet('Unidades amostrais', { landscape: true });
    r = U.title(0, 'Riqueza e abundância por unidade amostral', titulo + ' · ' + periodo, 8);
    U.width(0, 22); for (var c = 1; c <= Math.max(nU, 3) + 1; c++) U.width(c, 11);
    var hR = r; U.set(r, 0, '', S.th); D.uas.forEach(function (u, k) { U.set(r, k + 1, u.nome, S.th); }); U.set(r, nU + 1, 'Total', S.th); r++;
    var rRiq = r; U.set(r, 0, 'Riqueza', S.totl); D.uas.forEach(function (u, k) { U.set(r, k + 1, D.byUA[k].t.filter(Boolean).length, st('td', { nf: '0' })); }); U.set(r, nU + 1, nS, st('tot', { nf: '0' })); r++;
    var rAb = r; U.set(r, 0, 'Abundância', S.totl); D.uas.forEach(function (u, k) { U.set(r, k + 1, D.byUA[k].t.reduce(function (a, b) { return a + b; }, 0), st('td', { nf: '0' })); }); U.set(r, nU + 1, D.N, st('tot', { nf: '0' })); r++;
    U.set(r, 0, 'Listas (amostras)', S.totl); D.uas.forEach(function (u, k) { U.set(r, k + 1, u.listas.length, st('td', { nf: '0' })); }); U.set(r, nU + 1, nL, st('tot', { nf: '0' })); r++;
    var catsUA = { ref: absR(U.name, hR, 1, hR, nU), vals: D.uas.map(function (u) { return u.nome; }) };
    U.chart({ title: 'Riqueza e Abundância por Unidade Amostral', xTitle: 'Unidades amostrais', cats: catsUA, series: [
      { name: 'Riqueza', nameRef: absR(U.name, rRiq, 0), ref: absR(U.name, rRiq, 1, rRiq, nU), vals: D.uas.map(function (u, k) { return D.byUA[k].t.filter(Boolean).length; }), color: C.verde, labels: true },
      { name: 'Abundância', nameRef: absR(U.name, rAb, 0), ref: absR(U.name, rAb, 1, rAb, nU), vals: D.uas.map(function (u, k) { return D.byUA[k].t.reduce(function (a, b) { return a + b; }, 0); }), color: C.bronze, labels: true }],
      at: { c1: 0, r1: r + 1, c2: Math.max(nU + 2, 7), r2: r + 21 } });
    r += 23; U.pageBreak(r - 1);
    // métodos e tipo de registro
    var hM = r; U.set(r, 0, 'Método', S.th); U.set(r, 1, 'Riqueza', S.th); U.set(r, 2, 'Abundância', S.th); U.set(r, 3, 'Listas', S.th); r++;
    var rM0 = r;
    D.mets.forEach(function (m, k) { U.set(r, 0, m.nome, S.tdl); U.set(r, 1, D.byMet[k].filter(Boolean).length, st('td', { nf: '0' })); U.set(r, 2, D.byMet[k].reduce(function (a, b) { return a + b; }, 0), st('td', { nf: '0' })); U.set(r, 3, D.listas.filter(function (L) { return L.metodo === m.nome; }).length, st('td', { nf: '0' })); r++; });
    U.chart({ title: 'Riqueza e Abundância por Método', xTitle: 'Metodologia', cats: { ref: absR(U.name, rM0, 0, r - 1, 0), vals: D.mets.map(function (m) { return m.nome; }) }, series: [
      { name: 'Riqueza', nameRef: absR(U.name, hM, 1), ref: absR(U.name, rM0, 1, r - 1, 1), vals: D.mets.map(function (m, k) { return D.byMet[k].filter(Boolean).length; }), color: C.verde, labels: true },
      { name: 'Abundância', nameRef: absR(U.name, hM, 2), ref: absR(U.name, rM0, 2, r - 1, 2), vals: D.mets.map(function (m, k) { return D.byMet[k].reduce(function (a, b) { return a + b; }, 0); }), color: C.bronze, labels: true }],
      at: { c1: 0, r1: r + 1, c2: 5, r2: r + 19 } });
    var vT = D.sp.reduce(function (s, e) { return s + e.v; }, 0), aT = D.sp.reduce(function (s, e) { return s + e.a; }, 0), oT = D.N - vT - aT;
    var rV = hM; U.set(rV, 6, 'Tipo de registro', S.th); U.set(rV, 7, 'Riqueza', S.th); U.set(rV, 8, 'Abundância', S.th);
    var tipos = [['Visual', D.sp.filter(function (e) { return e.v; }).length, vT], ['Auditivo', D.sp.filter(function (e) { return e.a; }).length, aT]];
    if (oT > 0) tipos.push(['Sem tipo', D.sp.filter(function (e) { return e.o; }).length, oT]);
    tipos.forEach(function (t, k) { U.set(rV + 1 + k, 6, t[0], S.tdl); U.set(rV + 1 + k, 7, t[1], st('td', { nf: '0' })); U.set(rV + 1 + k, 8, t[2], st('td', { nf: '0' })); });
    U.width(6, 14); U.width(7, 10); U.width(8, 11);
    U.chart({ title: 'Método de Registro', cats: { ref: absR(U.name, rV + 1, 6, rV + tipos.length, 6), vals: tipos.map(function (t) { return t[0]; }) }, series: [
      { name: 'Riqueza', nameRef: absR(U.name, rV, 7), ref: absR(U.name, rV + 1, 7, rV + tipos.length, 7), vals: tipos.map(function (t) { return t[1]; }), color: C.verde, labels: true },
      { name: 'Abundância', nameRef: absR(U.name, rV, 8), ref: absR(U.name, rV + 1, 8, rV + tipos.length, 8), vals: tipos.map(function (t) { return t[2]; }), color: C.bronze, labels: true }],
      at: { c1: 6, r1: r + 1, c2: 11, r2: r + 19 } });

    /* ---------- 4. Riqueza acumulada (campanhas ou dias) ---------- */
    var A = B.sheet('Riqueza acumulada', { landscape: true });
    var useCamp = nC > 1, groups;
    if (useCamp) groups = D.camps.map(function (c, k) { return { nome: c.nome, vec: D.byCamp[k] }; });
    else {
      var byDay = {}, order = [];
      D.listas.forEach(function (L) { var dk = fmtDate(L.inicio); if (!byDay[dk]) { byDay[dk] = D.sp.map(function () { return 0; }); order.push(dk); } D.byList[L._i].forEach(function (v, i) { byDay[dk][i] += v; }); });
      groups = order.map(function (dk) { return { nome: dk, vec: byDay[dk] }; });
    }
    r = A.title(0, useCamp ? 'Riqueza acumulada por campanha' : 'Riqueza acumulada por dia de amostragem', titulo + ' · ' + periodo + (useCamp ? '' : ' · com mais de uma campanha na seleção, o acúmulo é feito por campanha.'), 12);
    A.width(0, 18); groups.forEach(function (g, k) { A.width(k + 1, Math.max(12, Math.min(20, String(g.nome).length + 2))); });
    var hA = r; A.set(r, 0, '', S.th); groups.forEach(function (g, k) { A.set(r, k + 1, g.nome, S.th); }); A.height(r, 24); r++;
    var seen = {}, acc = 0, porG = [], novas = [], acum = [];
    groups.forEach(function (g) { var s = 0, n = 0; g.vec.forEach(function (v, i) { if (v > 0) { s++; if (!seen[i]) { seen[i] = 1; n++; } } }); acc += n; porG.push(s); novas.push(n); acum.push(acc); });
    var rPor = r; A.set(r, 0, useCamp ? 'Por campanha' : 'Por dia', S.totl); porG.forEach(function (v, k) { A.set(r, k + 1, v, st('td', { nf: '0' })); }); r++;
    var rNov = r; A.set(r, 0, 'Espécies novas', S.totl); novas.forEach(function (v, k) { A.set(r, k + 1, v, st('td', { nf: '0' })); }); r++;
    var rAcc = r; A.set(r, 0, 'Acumulado', S.totl); acum.forEach(function (v, k) { A.set(r, k + 1, v, st('tot', { nf: '0' }), k === 0 ? null : 'SUM(' + A1(r, k) + ',' + A1(rNov, k + 1) + ')'); }); r++;
    var rAbG = r; A.set(r, 0, 'Abundância', S.totl); groups.forEach(function (g, k) { A.set(r, k + 1, g.vec.reduce(function (a, b) { return a + b; }, 0), st('td', { nf: '0' })); }); r++;
    var catsA = { ref: absR(A.name, hA, 1, hA, groups.length), vals: groups.map(function (g) { return g.nome; }) };
    A.chart({ title: useCamp ? 'Riqueza Acumulada por Campanha' : 'Riqueza Acumulada por Dia', xTitle: useCamp ? 'Campanhas de amostragem' : 'Dias de amostragem', yTitle: 'Riqueza', cats: catsA, series: [
      { name: useCamp ? 'Por campanha' : 'Por dia', nameRef: absR(A.name, rPor, 0), ref: absR(A.name, rPor, 1, rPor, groups.length), vals: porG, color: C.verde, labels: true, kind: 'bar' },
      { name: 'Acumulado', nameRef: absR(A.name, rAcc, 0), ref: absR(A.name, rAcc, 1, rAcc, groups.length), vals: acum, color: C.bronze, labels: true, kind: 'line', marker: 'circle', msize: 7 }],
      at: { c1: 0, r1: r + 1, c2: Math.max(groups.length + 2, 9), r2: r + 22 } });
    r += 24;
    // matriz de presença (cumulativa) estilo "Acum e Freq"
    A.set(r, 0, 'Presença por ' + (useCamp ? 'campanha' : 'dia') + ' — N: primeiro registro; X: registrada novamente; f: nº de ' + (useCamp ? 'campanhas' : 'dias') + ' com registro.', S.nota); A.merge(r, 0, r, Math.max(groups.length + 3, 8), S.nota); r++;
    var nG = groups.length, hP = r; A.pageBreak(r - 1); A.set(r, 0, 'Espécie', S.th); groups.forEach(function (g, k) { A.set(r, k + 1, g.nome, S.th); }); A.set(r, nG + 1, 'f', S.th); A.set(r, nG + 2, '%', S.th); A.height(r, 24); r++;
    A.width(0, 26); A.width(nG + 1, 6); A.width(nG + 2, 7);
    D.sp.forEach(function (e) {
      var first = true, f = 0; A.set(r, 0, e.nome, S.sp);
      groups.forEach(function (g, k) { var v = g.vec[e.i] > 0; if (v) { A.set(r, k + 1, first ? 'N' : 'X', S.td); first = false; f++; } else A.set(r, k + 1, null, S.td); });
      var rng2 = A1(r, 1) + ':' + A1(r, nG);
      A.set(r, nG + 1, f, st('tot', { nf: '0' }), 'COUNTIF(' + rng2 + ',"X")+COUNTIF(' + rng2 + ',"N")');
      A.set(r, nG + 2, round(100 * f / nG, 1), st('tot', { nf: '0.0' }), A1(r, nG + 1) + '*100/' + nG);
      r++;
    });

    /* ---------- 5. Diversidade ---------- */
    var V = B.sheet('Diversidade', { landscape: true });
    r = V.title(0, 'Índices de diversidade', titulo + ' · ' + periodo + ' · logaritmo natural (ln).', 10);
    V.width(0, 30); for (c = 1; c <= nU + 1; c++) V.width(c, 12);
    var hV = r; V.set(r, 0, 'Índice', S.th); D.uas.forEach(function (u, k) { V.set(r, k + 1, u.nome, S.th); }); V.set(r, nU + 1, 'Total', S.th); r++;
    var colsV = D.byUA.map(function (u) { return u.t; }).concat([totCounts]);
    var rowsV = [
      ['Riqueza (S)', function (v) { return v.filter(Boolean).length; }, '0'],
      ['Abundância (N)', function (v) { return v.reduce(function (a, b) { return a + b; }, 0); }, '0'],
      ['Shannon (H\')', function (v) { return round(shannon(v), 4); }, '0.000'],
      ['H máx = ln S', function (v) { var s = v.filter(Boolean).length; return s ? round(Math.log(s), 4) : 0; }, '0.000'],
      ['Equitabilidade de Pielou (J\')', function (v) { var s = v.filter(Boolean).length; return s > 1 ? round(shannon(v) / Math.log(s), 4) : null; }, '0.000'],
      ['Simpson (1 − D)', function (v) { var x = simpson(v); return x == null ? null : round(x, 4); }, '0.000'],
      ['Dominância de Berger-Parker', function (v) { var n = v.reduce(function (a, b) { return a + b; }, 0); return n ? round(Math.max.apply(null, v) / n, 4) : null; }, '0.000'],
      ['Riqueza de Margalef', function (v) { var n = v.reduce(function (a, b) { return a + b; }, 0), s = v.filter(Boolean).length; return n > 1 ? round((s - 1) / Math.log(n), 4) : null; }, '0.000'],
      ['Chao 1', function (v) { return round(chao1(v), 2); }, '0.00']];
    var rH = 0, rJ = 0;
    rowsV.forEach(function (x, j) { if (j === 2) rH = r; if (j === 4) rJ = r; V.set(r, 0, x[0], S.totl); colsV.forEach(function (v, k) { V.set(r, k + 1, x[1](v), st(k === nU ? 'tot' : 'td', { nf: x[2] })); }); r++; });
    V.chart({ title: 'Diversidade de Shannon (H\') e Equitabilidade (J\')', xTitle: 'Unidades amostrais', cats: { ref: absR(V.name, hV, 1, hV, nU), vals: D.uas.map(function (u) { return u.nome; }) }, series: [
      { name: 'Shannon (H\')', ref: absR(V.name, rH, 1, rH, nU), vals: colsV.slice(0, nU).map(function (v) { return round(shannon(v), 3); }), color: C.verde, labels: true },
      { name: 'Pielou (J\')', ref: absR(V.name, rJ, 1, rJ, nU), vals: colsV.slice(0, nU).map(function (v) { var s = v.filter(Boolean).length; return s > 1 ? round(shannon(v) / Math.log(s), 3) : 0; }), color: C.bronze, labels: true }],
      yFmt: '0.0', at: { c1: 0, r1: r + 1, c2: Math.max(nU + 2, 7), r2: r + 20 } });
    r += 22;
    V.set(r, 0, 'Cálculo de Shannon (total do levantamento): pi = ni/N; H\' = −Σ pi · ln pi.', S.nota); V.merge(r, 0, r, 6, S.nota); r++;
    var hPi = r; ['Espécie', 'ni', 'pi', 'ln pi', 'pi · ln pi'].forEach(function (h, i) { V.set(r, i, h, S.th); }); r++;
    var rPi0 = r;
    D.sp.forEach(function (e) {
      var p = D.N ? e.N / D.N : 0;
      V.set(r, 0, e.nome, S.sp); V.set(r, 1, e.N, st('td', { nf: '0' }));
      V.set(r, 2, round(p, 6), st('td', { nf: '0.0000' }), A1(r, 1) + '/' + A1(rPi0 + nS, 1));
      V.set(r, 3, p ? round(Math.log(p), 6) : null, st('td', { nf: '0.0000' }), 'IF(' + A1(r, 2) + '>0,LN(' + A1(r, 2) + '),0)');
      V.set(r, 4, p ? round(p * Math.log(p), 6) : 0, st('td', { nf: '0.0000' }), A1(r, 2) + '*' + A1(r, 3));
      r++;
    });
    V.set(r, 0, 'Total (N)', S.totl); V.set(r, 1, D.N, st('tot', { nf: '0' }), 'SUM(' + A1(rPi0, 1) + ':' + A1(r - 1, 1) + ')'); V.fill(r, 2, r, 3, S.tot); V.set(r, 4, round(-H, 6), st('tot', { nf: '0.0000' }), 'SUM(' + A1(rPi0, 4) + ':' + A1(r - 1, 4) + ')'); r++;
    V.set(r, 0, 'H\'', S.totl); V.set(r, 4, round(H, 4), st('tot', { nf: '0.0000' }), '-' + A1(r - 1, 4)); V.fill(r, 1, r, 3, S.tot); r++;
    V.set(r, 0, 'J\' = H\'/ln S', S.totl); V.set(r, 4, J == null ? null : round(J, 4), st('tot', { nf: '0.0000' }), nS > 1 ? A1(r - 1, 4) + '/LN(' + nS + ')' : null); V.fill(r, 1, r, 3, S.tot);
    V.width(1, 8); V.width(2, 10); V.width(3, 10); V.width(4, 11);

    /* ---------- 6. Rarefação e estimadores ---------- */
    var Rf = B.sheet('Rarefação', { landscape: true });
    r = Rf.title(0, 'Curva de rarefação e estimadores de riqueza', titulo + ' · amostras = listas (' + nL + ') · S(est): rarefação por amostras (Mao Tau, Colwell et al. 2004) com IC 95%; Jackknife 1: média de 100 aleatorizações.', 10);
    [10, 12, 12, 12, 10, 12, 10, 10, 12].forEach(function (w, i) { Rf.width(i, w); });
    if (nL >= 2) {
      var mt = maoTau(incAll), jc = jackCurve(D.byList, 100);
      var hRf = r; ['Amostras', 'S(est)', 'IC 95% inferior', 'IC 95% superior', 'S(est) DP', 'Jackknife 1', 'Jack 1 DP', 'IC Jack (±)', 'IC S(est) (±)'].forEach(function (h, i) { Rf.set(r, i, h, S.th); }); Rf.height(r, 26); r++;
      var rRf0 = r;
      mt.forEach(function (x, k) {
        Rf.set(r, 0, x.h, st('td', { nf: '0' })); Rf.set(r, 1, round(x.S, 2), st('td', { nf: '0.00' })); Rf.set(r, 2, round(x.lo, 2), st('td', { nf: '0.00' })); Rf.set(r, 3, round(x.hi, 2), st('td', { nf: '0.00' })); Rf.set(r, 4, round(x.sd, 2), st('td', { nf: '0.00' }));
        Rf.set(r, 5, round(jc[k].mean, 2), st('td', { nf: '0.00' })); Rf.set(r, 6, round(jc[k].sd, 2), st('td', { nf: '0.00' }));
        Rf.set(r, 7, round(1.96 * jc[k].sd, 2), st('td', { nf: '0.00' }), '1.96*' + A1(r, 6)); Rf.set(r, 8, round(1.96 * x.sd, 2), st('td', { nf: '0.00' }), '1.96*' + A1(r, 4));
        r++;
      });
      Rf.chart({ type: 'line', title: 'Curva de Rarefação e Jackknife 1', xTitle: 'Número de amostras (listas)', yTitle: 'Riqueza de espécies', skip: nL > 20 ? Math.ceil(nL / 15) : null,
        cats: { ref: absR(Rf.name, rRf0, 0, r - 1, 0), vals: mt.map(function (x) { return x.h; }), num: true }, series: [
          { name: 'Riqueza observada (rarefação)', nameRef: null, ref: absR(Rf.name, rRf0, 1, r - 1, 1), vals: mt.map(function (x) { return round(x.S, 2); }), color: C.verde, err: { ref: absR(Rf.name, rRf0, 8, r - 1, 8), vals: mt.map(function (x) { return round(1.96 * x.sd, 2); }) } },
          { name: 'Jackknife 1', ref: absR(Rf.name, rRf0, 5, r - 1, 5), vals: jc.map(function (x) { return round(x.mean, 2); }), color: C.bronze, marker: 'square', err: { ref: absR(Rf.name, rRf0, 7, r - 1, 7), vals: jc.map(function (x) { return round(1.96 * x.sd, 2); }) } }],
        at: { c1: 10, r1: hRf, c2: 19, r2: hRf + 22 } });
      r++;
    } else { Rf.set(r, 0, 'A curva de rarefação precisa de pelo menos 2 listas (amostras) na seleção.', S.nota); Rf.merge(r, 0, r, 8, S.nota); r += 2; }
    var hEs = r; ['Estimador', 'Valor', 'Completude (%)', 'Base'].forEach(function (h, i) { Rf.set(r, i === 3 ? 3 : i, h, S.th); }); Rf.merge(r, 3, r, 6, S.th); r++;
    var Q = incAll, ests = [['Riqueza observada (S obs)', nS, 'Espécies registradas'], ['Chao 1', chao1(totCounts), 'Abundância: singletons (' + totCounts.filter(function (n) { return n === 1; }).length + ') e doubletons (' + totCounts.filter(function (n) { return n === 2; }).length + ')'],
      ['Chao 2', nL > 1 ? chao2(Q) : null, 'Incidência: únicas (' + Q.Q1 + ') e duplicadas (' + Q.Q2 + ') em ' + nL + ' listas'], ['Jackknife 1', nL > 1 ? jack1(Q) : null, 'Incidência'], ['Jackknife 2', nL > 2 ? jack2(Q) : null, 'Incidência']];
    ests.forEach(function (x) { Rf.set(r, 0, x[0], S.tdl); Rf.set(r, 1, x[1] == null ? '—' : round(x[1], 2), st('td', { nf: '0.00' })); Rf.set(r, 2, x[1] ? round(100 * nS / x[1], 1) : '—', st('td', { nf: '0.0' })); Rf.set(r, 3, x[2], S.tdl); Rf.merge(r, 3, r, 6, S.tdl); r++; });
    Rf.width(0, 26);

    /* ---------- 7. Similaridade ---------- */
    var Sm = B.sheet('Similaridade', { landscape: true });
    r = Sm.title(0, 'Similaridade entre unidades amostrais', titulo + ' · Jaccard (presença/ausência) e Bray-Curtis (abundância), de 0 (nada em comum) a 1 (idênticas).', 10);
    Sm.width(0, 20); for (c = 1; c <= nU; c++) Sm.width(c, 11);
    [['Jaccard', jaccard, 'presença'], ['Bray-Curtis', brayCurtisSim, 'abundância']].forEach(function (M) {
      Sm.set(r, 0, 'Índice de ' + M[0], S.secao); Sm.fill(r, 1, r, nU, S.secao); r++;
      Sm.set(r, 0, '', S.th); D.uas.forEach(function (u, k) { Sm.set(r, k + 1, u.nome, S.th); }); r++;
      D.uas.forEach(function (u, i) {
        Sm.set(r, 0, u.nome, S.th2);
        D.uas.forEach(function (w, j) { var v = M[1](D.byUA[i].t, D.byUA[j].t); Sm.set(r, j + 1, v == null ? '—' : round(v, 3), st(i === j ? 'tot' : 'td', { nf: '0.000' })); });
        r++;
      });
      r++;
    });
    if (nC > 1) {
      Sm.set(r, 0, 'Similaridade de Jaccard entre campanhas', S.secao); Sm.fill(r, 1, r, nC, S.secao); r++;
      Sm.set(r, 0, '', S.th); D.camps.forEach(function (cc, k) { Sm.set(r, k + 1, cc.nome, S.th); Sm.width(k + 1, Math.max(Sm.cols[k + 1] || 0, 11)); }); r++;
      D.camps.forEach(function (cc, i) { Sm.set(r, 0, cc.nome, S.th2); D.camps.forEach(function (w, j) { var v = jaccard(D.byCamp[i], D.byCamp[j]); Sm.set(r, j + 1, v == null ? '—' : round(v, 3), st(i === j ? 'tot' : 'td', { nf: '0.000' })); }); r++; });
    }

    /* ---------- 8. IPA ---------- */
    var I = B.sheet('IPA', { landscape: true });
    r = I.title(0, 'Índice Pontual de Abundância (IPA)', titulo + ' · IPA = nº de indivíduos da espécie / nº de amostras (listas) da unidade amostral.', 10);
    I.width(0, 28); for (c = 1; c <= nU + 2; c++) I.width(c, 10);
    var hI = r; I.set(r, 0, 'Espécie', S.th); D.uas.forEach(function (u, k) { I.set(r, k + 1, u.nome + ' (' + u.listas.length + ')', S.th); }); I.set(r, nU + 1, 'Geral (' + nL + ')', S.th); r++;
    var ipaRows = D.sp.map(function (e) { return { e: e, g: nL ? e.N / nL : 0 }; }).sort(function (a, b) { return b.g - a.g; });
    var rI0 = r;
    ipaRows.forEach(function (x) { I.set(r, 0, x.e.nome, S.sp); D.uas.forEach(function (u, k) { var v = D.byUA[k].t[x.e.i]; I.set(r, k + 1, v ? round(v / u.listas.length, 3) : null, st('td', { nf: '0.000' })); }); I.set(r, nU + 1, round(x.g, 3), st('tot', { nf: '0.000' })); r++; });
    var top = Math.min(10, ipaRows.length);
    if (top) I.chart({ title: 'IPA — ' + top + ' espécies mais abundantes', cats: { ref: absR(I.name, rI0, 0, rI0 + top - 1, 0), vals: ipaRows.slice(0, top).map(function (x) { return x.e.nome; }) }, series: [
      { name: 'IPA geral', ref: absR(I.name, rI0, nU + 1, rI0 + top - 1, nU + 1), vals: ipaRows.slice(0, top).map(function (x) { return round(x.g, 2); }), color: C.verde, labels: true }], legend: false, yTitle: 'IPA', yFmt: '0.0',
      at: { c1: nU + 3, r1: hI, c2: nU + 12, r2: hI + 22 } });

    /* ---------- 9. Composição ---------- */
    var Cp = B.sheet('Composição', { landscape: true });
    r = Cp.title(0, 'Composição taxonômica e conservação', titulo + ' · ' + periodo, 10);
    function contagem(fn) { var m = {}; D.sp.forEach(function (e) { var k = fn(e) || '—'; if (!m[k]) m[k] = { s: 0, n: 0 }; m[k].s++; m[k].n += e.N; }); return Object.keys(m).map(function (k) { return { k: k, s: m[k].s, n: m[k].n }; }).sort(function (a, b) { return b.s - a.s || a.k.localeCompare(b.k); }); }
    var blocos = [['Ordem', contagem(function (e) { return e.d.o; }), true], ['Família', contagem(function (e) { return e.d.f; }), true], ['Endêmica do Brasil', contagem(function (e) { return e.d.e || 'Sem informação'; })], ['Origem', contagem(function (e) { return e.d.og || 'Sem informação'; })],
      ['IUCN', contagem(function (e) { return e.d.iucn ? e.d.iucn + ' – ' + (IUCN_NOME[e.d.iucn] || '') : 'Sem informação'; })], ['MMA', contagem(function (e) { return e.d.mma ? e.d.mma + ' – ' + (IUCN_NOME[e.d.mma] || '') : 'Sem informação'; })], ['Ocorrência (aves – CBRO)', contagem(function (e) { return e.d.c === 'Ave' ? e.d.oc || 'Sem informação' : null; }).filter(function (x) { return x.k !== '—'; })]];
    Cp.width(0, 30); Cp.width(1, 10); Cp.width(2, 12); Cp.width(3, 3); Cp.width(4, 30); Cp.width(5, 10); Cp.width(6, 12);
    var left = r, right = r, chartsAt = [];
    blocos.forEach(function (b, bi) {
      if (!b[1].length) return;
      var col = bi % 2 === 0 ? 0 : 4, rr = bi % 2 === 0 ? left : right;
      Cp.set(rr, col, b[0], S.th); Cp.set(rr, col + 1, 'Espécies', S.th); Cp.set(rr, col + 2, 'Indivíduos', S.th); var h0 = rr; rr++;
      var r0 = rr; b[1].forEach(function (x) { Cp.set(rr, col, x.k, b[0] === 'Ordem' || b[0] === 'Família' ? S.tdl : S.tdl); Cp.set(rr, col + 1, x.s, st('td', { nf: '0' })); Cp.set(rr, col + 2, x.n, st('td', { nf: '0' })); rr++; });
      if (b[2]) chartsAt.push({ name: b[0], h0: h0, r0: r0, r1: rr - 1, col: col, vals: b[1] });
      rr++;
      if (bi % 2 === 0) left = rr; else right = rr;
    });
    r = Math.max(left, right) + 1; Cp.pageBreak(r - 1);
    chartsAt.forEach(function (ch, k) {
      var n = ch.vals.length;
      Cp.chart({ title: 'Contagem de espécies por ' + ch.name, cats: { ref: absR(Cp.name, ch.r0, ch.col, ch.r1, ch.col), vals: ch.vals.map(function (x) { return x.k; }) }, series: [
        { name: 'Espécies', ref: absR(Cp.name, ch.r0, ch.col + 1, ch.r1, ch.col + 1), vals: ch.vals.map(function (x) { return x.s; }), color: k === 0 ? C.verde : C.bronze, labels: true }], legend: false, yTitle: 'Espécies', yInt: Math.max(1, Math.ceil(Math.max.apply(null, ch.vals.map(function (x) { return x.s; })) / 6)),
        at: { c1: k === 0 ? 0 : 4, r1: r, c2: k === 0 ? 3 : 8, r2: r + 20 } });
    });

    /* ---------- 10. Registros (dados brutos) ---------- */
    var Rg = B.sheet('Registros', { landscape: true, freeze: { r: 3, c: 0 }, printTitles: [2, 2] });
    r = Rg.title(0, 'Registros de campo', titulo + ' · um registro por espécie em cada lista.', 12);
    var headR = ['Data', 'Hora', 'Campanha', 'Unidade amostral', 'Método', 'Lista', 'Classe', 'Ordem', 'Família', 'Espécie', 'Autor', 'Visual', 'Auditivo', 'Sem tipo', 'Total', 'Latitude', 'Longitude', 'Observação'];
    headR.forEach(function (h, i) { Rg.set(r, i, h, S.th); }); Rg.height(r, 24); r++;
    D.registros.slice().sort(function (a, b) { return (a.r.hora || a.L.inicio) < (b.r.hora || b.L.inicio) ? -1 : 1; }).forEach(function (x) {
      var d = x.e.d, t = x.r.hora || x.L.inicio;
      var row = [dayOnly(t), timeOnly(t), x.L.campanha, x.L.ua, x.L.metodo, x.L.titulo || '', CLASSE_NOME[d.c] || '', d.o || '', d.f || '', x.e.nome, x.e.autor || '', x.r.v || 0, x.r.a || 0, x.r.outro || 0, x.tot, x.r.lat != null ? x.r.lat : x.L.lat, x.r.lon != null ? x.r.lon : x.L.lon, x.r.obs || ''];
      row.forEach(function (v, i) { Rg.set(r, i, v == null ? null : v, i === 0 ? st('td', { nf: 'dd/mm/yyyy' }) : i === 1 ? st('td', { nf: 'hh:mm' }) : i === 9 ? S.sp : i >= 11 && i <= 14 ? st('td', { nf: '0' }) : i === 15 || i === 16 ? st('td', { nf: '0.00000' }) : i === 17 || i === 5 ? S.tdl : S.td); });
      r++;
    });
    [11, 7, 14, 14, 14, 18, 10, 15, 16, 26, 22, 7, 8, 8, 7, 11, 11, 30].forEach(function (w, i) { Rg.width(i, w); });

    /* ---------- 11. EstimateS (matriz) ---------- */
    var Es = B.sheet('Matriz EstimateS', { landscape: true, freeze: { r: 3, c: 1 } });
    r = Es.title(0, 'Matriz espécies × amostras (abundância)', 'Pronta para EstimateS, PAST ou R (vegan). Cada coluna é uma lista do app.', 10);
    Es.width(0, 28);
    Es.set(r, 0, 'Espécie', S.th); D.listas.forEach(function (L, k) { Es.set(r, k + 1, fmtDate(L.inicio) + ' ' + (L.ua || ''), st('th', { rot: 90 })); Es.width(k + 1, 4.5); }); Es.height(r, 70); r++;
    D.sp.forEach(function (e) { Es.set(r, 0, e.nome, S.sp); D.listas.forEach(function (L, k) { Es.set(r, k + 1, D.byList[k][e.i], st('td', { nf: '0' })); }); r++; });

    /* ---------- 12. Pontos ---------- */
    var Pt = B.sheet('Pontos', { landscape: false });
    r = Pt.title(0, 'Unidades amostrais e coordenadas', 'Coordenadas da primeira lista com GPS em cada unidade · Datum WGS 84.', 7);
    ['Unidade amostral', 'Latitude', 'Longitude', 'Zona UTM', 'E (m)', 'N (m)', 'Listas'].forEach(function (h, i) { Pt.set(r, i, h, S.th); }); r++;
    D.uas.forEach(function (u) { var z = utm(u.lat, u.lon); Pt.set(r, 0, u.nome, S.tdl); Pt.set(r, 1, u.lat, st('td', { nf: '0.00000' })); Pt.set(r, 2, u.lon, st('td', { nf: '0.00000' })); Pt.set(r, 3, z ? z.zone + z.band : '—', S.td); Pt.set(r, 4, z ? z.e : null, st('td', { nf: '0' })); Pt.set(r, 5, z ? z.n : null, st('td', { nf: '0' })); Pt.set(r, 6, u.listas.length, st('td', { nf: '0' })); r++; });
    [20, 12, 12, 10, 11, 12, 8].forEach(function (w, i) { Pt.width(i, w); });
    r++;
    ['Lista', 'Data', 'Início', 'Unidade amostral', 'Latitude', 'Longitude', 'Distância (km)', 'Duração (min)'].forEach(function (h, i) { Pt.set(r, i, h, S.th); }); r++;
    D.listas.forEach(function (L) { [L.titulo || '', dayOnly(L.inicio), timeOnly(L.inicio), L.ua, L.lat, L.lon, L.distKm || null, L.durMin || null].forEach(function (v, i) { Pt.set(r, i, v, i === 1 ? st('td', { nf: 'dd/mm/yyyy' }) : i === 2 ? st('td', { nf: 'hh:mm' }) : i === 4 || i === 5 ? st('td', { nf: '0.00000' }) : i === 6 ? st('td', { nf: '0.00' }) : i === 0 ? S.tdl : S.td); }); r++; });
    Pt.width(7, 10);

    /* ---------- 13. Legenda e fontes ---------- */
    var Lg = B.sheet('Legenda e fontes', { landscape: false });
    r = Lg.title(0, 'Legenda, critérios e fontes', 'Lista de Campo · criado por Lucas Aøsf', 4);
    Lg.width(0, 26); Lg.width(1, 70);
    var leg = [
      ['Tipo de registro', 'V – visual; A – auditivo (vocalização). "Sem tipo": listas antigas, anteriores à marcação V/A.'],
      ['N.T', 'Número total de indivíduos registrados da espécie.'],
      ['Abundância relativa', 'Indivíduos da espécie ÷ total de indivíduos × 100.'],
      ['Estimativa de abundância', 'M – muito abundante (≥ 5%); A – abundante (1 a 5%); P – pouco abundante (< 1%).'],
      ['F.O. e constância', 'Frequência de ocorrência = listas com a espécie ÷ total de listas × 100. Constância (Dajoz 1983): constante ≥ 50%; acessória 25–50%; acidental < 25%.'],
      ['Shannon (H\')', 'H\' = −Σ pi · ln pi; Pielou J\' = H\'/ln S; Simpson 1 − D = 1 − Σ ni(ni−1)/N(N−1); Berger-Parker = Nmáx/N; Margalef = (S−1)/ln N.'],
      ['Estimadores', 'Chao 1 (abundância) e Chao 2, Jackknife 1 e 2 (incidência, listas como amostras), com correção de viés. Rarefação por amostras: Mao Tau (Colwell et al. 2004), IC 95% = ± 1,96 DP.'],
      ['Similaridade', 'Jaccard = a/(a+b+c); Bray-Curtis = 1 − Σ|xi − yi| / Σ(xi + yi).'],
      ['IPA', 'Índice pontual de abundância = indivíduos ÷ número de amostras (listas).'],
      ['IUCN / MMA', 'LC – pouco preocupante; NT – quase ameaçada; VU – vulnerável; EN – em perigo; CR – criticamente em perigo; DD – dados insuficientes; N/A – não avaliada. Disponível para mamíferos (SBMz).'],
      ['CITES', 'Apêndices I, II ou III (mamíferos, SBMz).'],
      ['Dieta (SBMz)', 'Ca – carnívoro; Fo – folívoro; Fr – frugívoro; Go – gomívoro; Gr – granívoro; Hb – herbívoro pastador; In – insetívoro; Nec – nectarívoro; On – onívoro; Ps – piscívoro; Se – predador de sementes; Hm – hematófago; Mi – mirmecófago.'],
      ['Locomoção (SBMz)', 'Aq – aquático; Ar – arborícola; Fs – fossorial; SA – semiaquático; Sc – escansorial; SF – semifossorial; Te – terrestre; Vo – voador.'],
      ['Ocorrência (aves)', 'CBRO: Residente – residente ou migrante reprodutivo; Visitante – visitante sazonal não reprodutivo (N: norte; S: sul; E: leste; W: oeste); Vagante – ocorrência irregular; presumido – status não confirmado.'],
      ['UF de ocorrência', 'Mamíferos: SBMz. Répteis: Costa et al. (2026), tabela 3.'],
      ['Espécies sem classe', 'Nomes escritos livremente no app e ainda não ligados a uma espécie do banco. No app, toque no nome e use "Ligar a uma espécie".']];
    leg.forEach(function (x) { Lg.set(r, 0, x[0], S.totl); Lg.set(r, 1, x[1], st('tdl', { wrap: true })); Lg.height(r, Math.max(15, 13 * Math.ceil(x[1].length / 85))); r++; });
    r++; Lg.set(r, 0, 'Fontes das listas de espécies', S.secao); Lg.set(r, 1, null, S.secao); r++;
    var F = input.fontes || {};
    Object.keys(F).forEach(function (k) { Lg.set(r, 0, k, S.totl); Lg.set(r, 1, F[k], st('tdl', { wrap: true })); Lg.height(r, 28); r++; });
    r++; Lg.set(r, 0, 'Referências dos métodos', S.secao); Lg.set(r, 1, null, S.secao); r++;
    ['Colwell R.K., Mao C.X. & Chang J. 2004. Interpolating, extrapolating, and comparing incidence-based species accumulation curves. Ecology 85: 2717–2727.', 'Chao A. 1987. Estimating the population size for capture-recapture data with unequal catchability. Biometrics 43: 783–791.', 'Burnham K.P. & Overton W.S. 1979. Robust estimation of population size when capture probabilities vary among animals. Ecology 60: 927–936.', 'Dajoz R. 1983. Ecologia geral. Petrópolis: Vozes.', 'Magurran A.E. 2004. Measuring biological diversity. Oxford: Blackwell.'].forEach(function (t) { Lg.set(r, 0, null, S.totl); Lg.set(r, 1, t, st('tdl', { wrap: true })); Lg.height(r, 28); r++; });

    return B.build();
  }

  /* ===================== CSV ===================== */
  function csvCell(v) { if (v == null) return ''; if (typeof v === 'number') return String(Math.round(v * 1e6) / 1e6).replace('.', ','); var s = String(v); return /[;"\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function csvRows(rows) { return '﻿' + rows.map(function (r) { return r.map(csvCell).join(';'); }).join('\r\n') + '\r\n'; }
  function csvRegistros(input) {
    var D = preparar(input), rows = [['Data', 'Hora', 'Campanha', 'Unidade amostral', 'Método', 'Lista', 'Classe', 'Ordem', 'Família', 'Espécie', 'Autor', 'Visual', 'Auditivo', 'Sem tipo', 'Total', 'Latitude', 'Longitude', 'Observação']];
    D.registros.forEach(function (x) { var d = x.e.d, t = x.r.hora || x.L.inicio, dt = new Date(t); rows.push([fmtDate(t), pad(dt.getHours()) + ':' + pad(dt.getMinutes()), x.L.campanha, x.L.ua, x.L.metodo, x.L.titulo || '', CLASSE_NOME[d.c] || '', d.o || '', d.f || '', x.e.nome, x.e.autor || '', x.r.v || 0, x.r.a || 0, x.r.outro || 0, x.tot, x.r.lat != null ? x.r.lat : x.L.lat, x.r.lon != null ? x.r.lon : x.L.lon, x.r.obs || '']); });
    return csvRows(rows);
  }
  function csvEspecies(input) {
    var D = preparar(input), nL = D.listas.length, rows = [['Classe', 'Ordem', 'Família', 'Espécie', 'Autor', 'Nome popular', 'Tipo de registro'].concat(D.uas.map(function (u) { return u.nome + ' (V)'; }), D.uas.map(function (u) { return u.nome + ' (A)'; }), ['N.T', 'Abund. rel. (%)', 'Estimativa', 'F.O. (%)', 'Constância', 'Origem', 'Endêmica BR', 'Ocorrência (aves)', 'IUCN', 'MMA', 'CITES', 'Dieta', 'Locomoção', 'UF de ocorrência', 'Biomas'])];
    D.sp.forEach(function (e) {
      var d = e.d, pct = D.N ? 100 * e.N / D.N : 0, fo = nL ? 100 * D.byList.filter(function (l) { return l[e.i] > 0; }).length / nL : 0;
      rows.push([CLASSE_NOME[d.c] || '', d.o || '', d.f || '', e.nome, e.autor || '', d.pt ? d.pt.split(',')[0].trim() : '', tipoReg(e)].concat(D.uas.map(function (u, k) { return D.byUA[k].v[e.i]; }), D.uas.map(function (u, k) { return D.byUA[k].t[e.i] - D.byUA[k].v[e.i]; }), [e.N, round(pct, 2), estimativa(pct), round(fo, 1), constancia(fo), d.og || '', d.e || '', d.oc || '', d.iucn || '', d.mma || '', d.cites || '', d.die || '', d.loc || '', d.uf || '', d.bio || '']));
    });
    return csvRows(rows);
  }
  function csvMatriz(input) { var D = preparar(input), rows = [['Espécie'].concat(D.listas.map(function (L) { return fmtDate(L.inicio) + ' ' + L.ua; }))]; D.sp.forEach(function (e) { rows.push([e.nome].concat(D.listas.map(function (L, k) { return D.byList[k][e.i]; }))); }); return csvRows(rows); }
  function csvZip(input, base) { return zip([{ name: base + '-registros.csv', data: csvRegistros(input) }, { name: base + '-especies.csv', data: csvEspecies(input) }, { name: base + '-matriz-estimates.csv', data: csvMatriz(input) }]); }

  var API = { gerar: gerar, csvRegistros: csvRegistros, csvEspecies: csvEspecies, csvMatriz: csvMatriz, csvZip: csvZip, preparar: preparar, utm: utm, _t: { maoTau: maoTau, incStats: incStats, chao2: chao2, jack1: jack1, jack2: jack2, chao1: chao1, shannon: shannon, simpson: simpson } };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else root.Planilha = API;
})(typeof window !== 'undefined' ? window : this);
