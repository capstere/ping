/* ClickLess Offline Analyzer v6
   - ES5 only (no const/let/arrow/template/module/async)
   - Offline XLSX reader: ZIP + raw DEFLATE (inflateRaw)
   - Reads: Seal Test NEG/POS (.xlsx), Worksheet (.xlsx), Test Summary (.csv)
   - Shows results on the page (no downloads).
*/

(function () {
  'use strict';

  // ----------------------------
  // DOM helpers
  // ----------------------------
  function $(id) { return document.getElementById(id); }

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    if (attrs) {
      for (var k in attrs) {
        if (!attrs.hasOwnProperty(k)) continue;
        if (k === 'text') node.textContent = String(attrs[k]);
        else if (k === 'html') node.innerHTML = String(attrs[k]);
        else if (k === 'class') node.className = String(attrs[k]);
        else if (k === 'style') node.setAttribute('style', String(attrs[k]));
        else node.setAttribute(k, String(attrs[k]));
      }
    }
    if (children && children.length) {
      for (var i = 0; i < children.length; i++) node.appendChild(children[i]);
    }
    return node;
  }

  function clearNode(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  function pill(text, kind) {
    var cls = 'pill';
    if (kind === 'ok') cls += ' ok';
    else if (kind === 'bad') cls += ' bad';
    else if (kind === 'warn') cls += ' warn';
    return el('span', { 'class': cls, text: text });
  }

  function setStatus(msg, isError) {
    var s = $('status');
    s.textContent = msg || '';
    s.className = 'small ' + (isError ? 'err' : 'muted');
  }

  // ----------------------------
  // File registry + detection
  // ----------------------------
  var TYPE_OPTIONS = [
    { key: 'auto', label: 'Auto' },
    { key: 'sealNeg', label: 'Seal Test NEG (xlsx)' },
    { key: 'sealPos', label: 'Seal Test POS (xlsx)' },
    { key: 'worksheet', label: 'Worksheet (xlsx)' },
    { key: 'testSummary', label: 'Test Summary (csv)' }
  ];

  var files = []; // {file, typeKey, detectedKey}

  function detectTypeByName(name) {
    var n = (name || '').toLowerCase();
    if (n.indexOf('seal test') >= 0 && (n.indexOf('neg') >= 0 || n.indexOf('negative') >= 0)) return 'sealNeg';
    if (n.indexOf('seal test') >= 0 && (n.indexOf('pos') >= 0 || n.indexOf('positive') >= 0)) return 'sealPos';
    if (n.indexOf('worksheet') >= 0) return 'worksheet';
    if (n.indexOf('test summary') >= 0 || n.indexOf('tests summary') >= 0 || n.slice(-4) === '.csv') return 'testSummary';
    if (n.slice(-4) === '.csv') return 'testSummary';
    if (n.slice(-5) === '.xlsx') return 'auto';
    return 'auto';
  }

  // ----------------------------
  // Zip bundle support (Alternative 2)
  // Drop 1 .zip containing the 4 required files.
  // Requires JSZip (loaded via index.html)
  // ----------------------------
  function isZipFile(file) {
    if (!file) return false;
    var n = (file.name || '').toLowerCase();
    if (n.slice(-4) === '.zip') return true;
    var t = (file.type || '').toLowerCase();
    if (t.indexOf('zip') >= 0) return true;
    return false;
  }

  function makeFileLike(data, name, mime) {
    var opts = mime ? { type: mime } : {};
    try {
      return new File([data], name, opts);
    } catch (e) {
      var b = new Blob([data], opts);
      b.name = name;
      return b;
    }
  }

  function readAsArrayBuffer(file) {
    // Prefer modern File/Blob.arrayBuffer, fallback to FileReader
    if (file && typeof file.arrayBuffer === 'function') {
      return file.arrayBuffer();
    }
    return new Promise(function (resolve, reject) {
      try {
        var fr = new FileReader();
        fr.onload = function () { resolve(fr.result); };
        fr.onerror = function () { reject(fr.error || new Error('FileReader error')); };
        fr.readAsArrayBuffer(file);
      } catch (e) {
        reject(e);
      }
    });
  }

  function extractBundleFilesFromZip(zipFile) {
    if (typeof JSZip === 'undefined') {
      return Promise.reject(new Error('JSZip not loaded'));
    }

    return readAsArrayBuffer(zipFile).then(function (buf) {
      return JSZip.loadAsync(buf);
    }).then(function (zip) {
      var picked = { sealNeg: null, sealPos: null, worksheet: null, testSummary: null };

      zip.forEach(function (relPath, entry) {
        if (!entry || entry.dir) return;
        if (relPath.indexOf('__MACOSX/') === 0) return;

        var base = relPath.split('/').pop();
        if (!base) return;
        if (base.indexOf('._') === 0) return;

        var det = detectTypeByName(base);
        if (!picked[det]) {
          picked[det] = { base: base, entry: entry };
        }
      });

      var out = [];
      var tasks = [];

      function pushPicked(key, mime) {
        if (!picked[key]) return;
        tasks.push(
          picked[key].entry.async('arraybuffer').then(function (ab) {
            out.push(makeFileLike(ab, picked[key].base, mime));
          })
        );
      }

      pushPicked('sealNeg', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      pushPicked('sealPos', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      pushPicked('worksheet', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      pushPicked('testSummary', 'text/csv');

      return Promise.all(tasks).then(function () { return out; });
    });
  }


  function typeLabel(key) {
    for (var i = 0; i < TYPE_OPTIONS.length; i++) if (TYPE_OPTIONS[i].key === key) return TYPE_OPTIONS[i].label;
    return key;
  }

  function refreshFileList() {
    var list = $('fileList');
    clearNode(list);

    if (!files.length) {
      list.appendChild(el('div', { 'class': 'small muted', text: 'Inga filer valda.' }));
      $('runBtn').disabled = true;
      return;
    }

    var tbl = el('table');
    var thead = el('thead');
    var trh = el('tr');
    trh.appendChild(el('th', { text: 'Fil' }));
    trh.appendChild(el('th', { text: 'Detekterad typ' }));
    trh.appendChild(el('th', { text: 'Anvand som' }));
    trh.appendChild(el('th', { text: '' }));
    thead.appendChild(trh);
    tbl.appendChild(thead);

    var tbody = el('tbody');
    for (var i = 0; i < files.length; i++) {
      (function (idx) {
        var f = files[idx];
        var tr = el('tr');

        tr.appendChild(el('td', { 'class': 'mono', text: f.file.name }));
        tr.appendChild(el('td', { text: typeLabel(f.detectedKey) }));

        var sel = el('select');
        for (var j = 0; j < TYPE_OPTIONS.length; j++) {
          var opt = el('option', { value: TYPE_OPTIONS[j].key, text: TYPE_OPTIONS[j].label });
          if (TYPE_OPTIONS[j].key === f.typeKey) opt.selected = true;
          sel.appendChild(opt);
        }
        sel.addEventListener('change', function () { files[idx].typeKey = sel.value; validateReady(); });
        tr.appendChild(el('td', null, [sel]));

        var rm = el('button', { 'class': 'btn2', type: 'button', text: 'Ta bort' });
        rm.addEventListener('click', function () { files.splice(idx, 1); refreshFileList(); validateReady(); });
        tr.appendChild(el('td', null, [rm]));

        tbody.appendChild(tr);
      })(i);
    }
    tbl.appendChild(tbody);
    list.appendChild(tbl);

    validateReady();
  }

    function validateReady() {
    var want = { sealNeg: 0, sealPos: 0, worksheet: 0, testSummary: 0 };
    for (var i = 0; i < files.length; i++) {
      var tk = files[i].typeKey;
      if (tk === 'auto') tk = files[i].detectedKey;
      if (want.hasOwnProperty(tk)) want[tk]++;
    }

    // v7: CSV validation is the first milestone — only require Test Summary for now.
    var okCsv = (want.testSummary >= 1);
    $('runBtn').disabled = !okCsv;

    if (!okCsv) {
      setStatus('Välj minst 1 fil: Test Summary (csv).', false);
      return;
    }

    // If other files are missing, we still run CSV-only and show a note.
    var missing = [];
    if (want.sealNeg < 1) missing.push('Seal NEG');
    if (want.sealPos < 1) missing.push('Seal POS');
    if (want.worksheet < 1) missing.push('Worksheet');

    if (missing.length) {
      setStatus('Redo för CSV-validering. (XLSX-delar saknas: ' + missing.join(', ') + ')', false);
    } else {
      setStatus('Redo. Klicka "Kör analys".', false);
    }
  }


  function addFiles(fileList) {
    if (!fileList || !fileList.length) return;

    var incoming = [];
    var tasks = [];

    for (var i = 0; i < fileList.length; i++) {
      var f = fileList[i];
      if (!f) continue;

      if (isZipFile(f)) {
        (function (zipF) {
          setStatus('Reading zip bundle: ' + (zipF.name || 'bundle.zip') + ' ...', false);
          tasks.push(
            extractBundleFilesFromZip(zipF).then(function (extracted) {
              if (extracted && extracted.length) {
                for (var j = 0; j < extracted.length; j++) incoming.push(extracted[j]);
              } else {
                setStatus('Zip did not contain recognizable files: ' + (zipF.name || ''), true);
              }
            }).catch(function (e) {
              var msg = (e && e.message) ? e.message : String(e);
              setStatus('Could not read zip: ' + (zipF.name || '') + ' (' + msg + ')', true);
            })
          );
        })(f);
        continue;
      }

      incoming.push(f);
    }

    Promise.all(tasks).then(function () {
      for (var k = 0; k < incoming.length; k++) {
        var ff = incoming[k];
        var detected = detectTypeByName(ff.name || '');
        files.push({ file: ff, typeKey: 'auto', detectedKey: detected });
      }

      refreshFileList();
      setStatus('', false);
    });
  }

  // ----------------------------
  // CSV parse
  // ----------------------------
    // ----------------------------
  // CSV parse + validation (Test Summary)
  // ----------------------------
  function parseCsv(text) {
    text = text || '';
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

    var lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

    // Guess delimiter from first non-empty line
    var delim = ';';
    for (var i = 0; i < lines.length; i++) {
      if (lines[i] && lines[i].length) {
        var cSemi = countChar(lines[i], ';');
        var cComma = countChar(lines[i], ',');
        delim = (cSemi >= cComma) ? ';' : ',';
        break;
      }
    }

    // Find header row dynamically (same logic as PowerShell: look for core columns)
    var headerIdx = -1;
    for (i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (!l) continue;
      var low = l.toLowerCase();
      if (low.indexOf('assay') >= 0 && low.indexOf('sample') >= 0 && (low.indexOf('cartridge') >= 0 || low.indexOf('s/n') >= 0)) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx < 0) return { ok: false, error: 'Kunde inte hitta header-rad i CSV.' };

    var headers = splitCsvLine(lines[headerIdx], delim);
    var rows = [];
    for (i = headerIdx + 1; i < lines.length; i++) {
      if (!lines[i]) continue;
      if (/^(\s*;)+\s*$/.test(lines[i]) || /^(\s*,)+\s*$/.test(lines[i])) continue;
      var cols = splitCsvLine(lines[i], delim);
      if (cols.length < 4) continue;
      rows.push(cols);
    }

    function safeCol(cols, idx) {
      if (idx < 0) return '';
      var v = (idx < cols.length) ? cols[idx] : '';
      if (v === null || typeof v === 'undefined') return '';
      return String(v);
    }

    function parseSampleId(sampleId) {
      sampleId = String(sampleId || '');
      var parts = sampleId.split('_');
      var bag = null;
      var ctrl = '';
      var pos = '';
      if (parts.length >= 2) {
        var bn = parseInt(parts[1], 10);
        if (!isNaN(bn)) bag = bn;
      }
      if (parts.length >= 3) ctrl = String(parts[2] || '');
      if (parts.length >= 4) pos = String(parts[3] || '');
      return { raw: sampleId, prefix: String(parts[0] || ''), bag: bag, ctrl: ctrl, pos: pos };
    }

    function classifyDetected(testResult) {
      var low = String(testResult || '').toLowerCase();
      if (!low) return 'unknown';
      if (low.indexOf('not detected') >= 0) return 'not_detected';
      if (low.indexOf('detected') >= 0) return 'detected';
      return 'unknown';
    }

    var idxAssay = findHeader(headers, /^assay$/i);
    var idxAssayVer = findHeader(headers, /^assay version$/i);
    var idxSample = findHeader(headers, /^sample id$/i);
    var idxCart = findHeader(headers, /^cartridge/i);
    var idxLot = findHeader(headers, /reagent lot id/i);
    var idxTestType = findHeader(headers, /^test type$/i);
    var idxTestResult = findHeader(headers, /^test result$/i);

    var countsSample = {};
    var countsCart = {};
    var assay = '';
    var assayVer = '';
    var lot = '';

    var bagCounts = {};
    var ctrlCounts = {}; // ctrl -> {detected, not_detected, unknown, total}
    var ctrlPrefixMap = {}; // ctrl -> {prefix -> count}
    var invalidRows = [];
    var liteRows = [];

    function incMap(map, key) { map[key] = (map[key] || 0) + 1; }

    for (i = 0; i < rows.length; i++) {
      var r = rows[i];

      var sAssay = safeCol(r, idxAssay);
      var sAssayVer = safeCol(r, idxAssayVer);
      var sSample = safeCol(r, idxSample);
      var sCart = safeCol(r, idxCart);
      var sLot = safeCol(r, idxLot);
      var sType = safeCol(r, idxTestType);
      var sRes = safeCol(r, idxTestResult);

      if (!assay && sAssay) assay = sAssay;
      if (!assayVer && sAssayVer) assayVer = sAssayVer;
      if (!lot && sLot) lot = sLot;

      if (sSample) incMap(countsSample, sSample);
      if (sCart) incMap(countsCart, sCart);

      var sid = parseSampleId(sSample);
      if (sid.bag !== null) incMap(bagCounts, String(sid.bag));

      var detClass = classifyDetected(sRes);
      if (!ctrlCounts.hasOwnProperty(sid.ctrl)) ctrlCounts[sid.ctrl] = { detected: 0, not_detected: 0, unknown: 0, total: 0 };
      ctrlCounts[sid.ctrl][detClass] = (ctrlCounts[sid.ctrl][detClass] || 0) + 1;
      ctrlCounts[sid.ctrl].total++;

      if (!ctrlPrefixMap.hasOwnProperty(sid.ctrl)) ctrlPrefixMap[sid.ctrl] = {};
      incMap(ctrlPrefixMap[sid.ctrl], sid.prefix);

      var lowRes = String(sRes || '').toLowerCase();
      var isInvalid = (lowRes.indexOf('no result') >= 0 ||
                       lowRes.indexOf('invalid') >= 0 ||
                       lowRes.indexOf('error') >= 0 ||
                       lowRes.indexOf('cancel') >= 0 ||
                       lowRes.indexOf('abort') >= 0);

      var lite = {
        assay: sAssay, assayVer: sAssayVer, lot: sLot,
        sampleId: sSample, cartSn: sCart,
        testType: sType, testResult: sRes,
        bag: sid.bag, ctrl: sid.ctrl, prefix: sid.prefix, pos: sid.pos,
        detectedClass: detClass
      };
      liteRows.push(lite);

      if (isInvalid) invalidRows.push(lite);
    }

    function makeDuplicates(counts) {
      var out = [];
      for (var k in counts) {
        if (!counts.hasOwnProperty(k)) continue;
        if (counts[k] > 1) out.push({ value: k, count: counts[k] });
      }
      out.sort(function (a, b) { return b.count - a.count; });
      return out;
    }

    // Bag range + missing (from min..max)
    var bagNums = [];
    for (var bk in bagCounts) if (bagCounts.hasOwnProperty(bk)) {
      var bn2 = parseInt(bk, 10);
      if (!isNaN(bn2)) bagNums.push(bn2);
    }
    bagNums.sort(function (a, b) { return a - b; });
    var bagMin = (bagNums.length ? bagNums[0] : null);
    var bagMax = (bagNums.length ? bagNums[bagNums.length - 1] : null);
    var bagMissing = [];
    if (bagMin !== null && bagMax !== null) {
      for (var b = bagMin; b <= bagMax; b++) {
        if (!bagCounts.hasOwnProperty(String(b))) bagMissing.push(b);
      }
    }

    // Control expected behavior = majority (detected vs not_detected). Unknown if tie/empty.
    var ctrlExpected = {}; // ctrl -> 'detected'|'not_detected'|'unknown'
    for (var ck in ctrlCounts) if (ctrlCounts.hasOwnProperty(ck)) {
      var c = ctrlCounts[ck];
      var d = c.detected || 0;
      var nd = c.not_detected || 0;
      if (d === 0 && nd === 0) ctrlExpected[ck] = 'unknown';
      else if (d > nd) ctrlExpected[ck] = 'detected';
      else if (nd > d) ctrlExpected[ck] = 'not_detected';
      else ctrlExpected[ck] = 'unknown';
    }

    // Deviations: any row that doesn't match expected (or is unknown) when expected is known
    var ctrlDeviations = []; // {ctrl, expected, sampleId, cartSn, testResult, detectedClass}
    for (i = 0; i < liteRows.length; i++) {
      var row = liteRows[i];
      var exp = ctrlExpected[row.ctrl] || 'unknown';
      if (exp === 'unknown') continue;
      if (row.detectedClass !== exp) {
        ctrlDeviations.push({
          ctrl: row.ctrl, expected: exp,
          sampleId: row.sampleId, cartSn: row.cartSn,
          testType: row.testType, testResult: row.testResult,
          detectedClass: row.detectedClass, prefix: row.prefix, bag: row.bag, pos: row.pos
        });
      }
    }

    return {
      ok: true,
      delim: delim,
      headerIdx: headerIdx,
      rowCount: rows.length,
      assay: assay,
      assayVer: assayVer,
      lot: lot,
      duplicateSamples: makeDuplicates(countsSample),
      duplicateCarts: makeDuplicates(countsCart),
      bag: { min: bagMin, max: bagMax, missing: bagMissing, counts: bagCounts },
      ctrl: { counts: ctrlCounts, expected: ctrlExpected, deviations: ctrlDeviations, prefixes: ctrlPrefixMap },
      invalid: invalidRows,
      // keep some raw for future rules
      _rows: liteRows
    };
  }


  function countChar(s, ch) {
    var n = 0;
    for (var i = 0; i < s.length; i++) if (s.charAt(i) === ch) n++;
    return n;
  }

  function splitCsvLine(line, delim) {
    var out = [];
    var cur = '';
    var inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line.charAt(i);
      if (inQ) {
        if (ch === '"') {
          if (i + 1 < line.length && line.charAt(i + 1) === '"') { cur += '"'; i++; }
          else inQ = false;
        } else {
          cur += ch;
        }
      } else {
        if (ch === '"') inQ = true;
        else if (ch === delim) { out.push(cur); cur = ''; }
        else cur += ch;
      }
    }
    out.push(cur);
    return out;
  }

  function findHeader(headers, regex) {
    for (var i = 0; i < headers.length; i++) {
      var h = (headers[i] || '').toString().trim();
      if (regex.test(h)) return i;
    }
    return -1;
  }

  function pickDuplicates(map) {
    var out = [];
    for (var k in map) {
      if (!map.hasOwnProperty(k)) continue;
      if (map[k] > 1) out.push({ key: k, count: map[k] });
    }
    out.sort(function (a, b) { return b.count - a.count; });
    return out;
  }

  // ----------------------------
  // ZIP + DEFLATE (raw) for XLSX
  // ----------------------------
  function readU16LE(bytes, off) { return bytes[off] | (bytes[off + 1] << 8); }
  function readU32LE(bytes, off) { return (bytes[off] | (bytes[off + 1] << 8) | (bytes[off + 2] << 16) | (bytes[off + 3] << 24)) >>> 0; }

  function decodeUtf8(bytes) {
    if (window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i] & 0xFF);
    return s;
  }

  function unzipLocal(bytes) {
    var out = {};
    var off = 0;
    while (off + 30 < bytes.length) {
      var sig = readU32LE(bytes, off);
      if (sig !== 0x04034b50) break;

      var flags = readU16LE(bytes, off + 6);
      var method = readU16LE(bytes, off + 8);
      var compSize = readU32LE(bytes, off + 18);
      var uncompSize = readU32LE(bytes, off + 22);
      var nameLen = readU16LE(bytes, off + 26);
      var extraLen = readU16LE(bytes, off + 28);

      var nameBytes = bytes.subarray(off + 30, off + 30 + nameLen);
      var name = decodeUtf8(nameBytes);

      var dataStart = off + 30 + nameLen + extraLen;
      var dataEnd = dataStart + compSize;
      if (dataEnd > bytes.length) break;

      var comp = bytes.subarray(dataStart, dataEnd);
      var data;
      if (method === 0) data = comp;
      else if (method === 8) data = inflateRaw(comp, uncompSize);
      else data = new Uint8Array(0);

      out[name] = data;

      off = dataEnd;
      if (flags & 0x08) break;
    }
    return out;
  }

  function BitReader(bytes) {
    this.bytes = bytes;
    this.pos = 0;
    this.bitbuf = 0;
    this.bitcnt = 0;
  }

  BitReader.prototype.readBits = function (n) {
    while (this.bitcnt < n) {
      if (this.pos >= this.bytes.length) throw new Error('Unexpected EOF in bitstream');
      this.bitbuf |= (this.bytes[this.pos++] << this.bitcnt);
      this.bitcnt += 8;
    }
    var val = this.bitbuf & ((1 << n) - 1);
    this.bitbuf >>>= n;
    this.bitcnt -= n;
    return val;
  };

  BitReader.prototype.alignByte = function () {
    this.bitbuf = 0;
    this.bitcnt = 0;
  };

  function buildHuffman(codeLengths) {
    var maxLen = 0;
    for (var i = 0; i < codeLengths.length; i++) if (codeLengths[i] > maxLen) maxLen = codeLengths[i];

    var blCount = new Array(maxLen + 1);
    for (i = 0; i < blCount.length; i++) blCount[i] = 0;

    for (i = 0; i < codeLengths.length; i++) {
      var len = codeLengths[i];
      if (len > 0) blCount[len]++;
    }

    var nextCode = new Array(maxLen + 1);
    var code = 0;
    blCount[0] = 0;
    for (i = 1; i <= maxLen; i++) {
      code = (code + blCount[i - 1]) << 1;
      nextCode[i] = code;
    }

    var root = {};
    for (var sym = 0; sym < codeLengths.length; sym++) {
      var l = codeLengths[sym];
      if (!l) continue;
      var c = nextCode[l]++;
      var node = root;
      for (var bit = l - 1; bit >= 0; bit--) {
        var b = (c >> bit) & 1;
        if (!node[b]) node[b] = {};
        node = node[b];
      }
      node.sym = sym;
    }
    return root;
  }

  function decodeSym(br, tree) {
    var node = tree;
    while (true) {
      var b = br.readBits(1);
      node = node[b];
      if (!node) throw new Error('Bad Huffman code');
      if (node.sym !== undefined) return node.sym;
    }
  }

  var LEN_BASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
  var LEN_EXTRA= [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  var DIST_BASE=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
  var DIST_EXTRA=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];

  function fixedLitLenTree() {
    var lengths = new Array(288);
    for (var i = 0; i <= 143; i++) lengths[i] = 8;
    for (i = 144; i <= 255; i++) lengths[i] = 9;
    for (i = 256; i <= 279; i++) lengths[i] = 7;
    for (i = 280; i <= 287; i++) lengths[i] = 8;
    return buildHuffman(lengths);
  }

  function fixedDistTree() {
    var lengths = new Array(32);
    for (var i = 0; i < 32; i++) lengths[i] = 5;
    return buildHuffman(lengths);
  }

  function inflateRaw(compBytes, expectedSize) {
    var br = new BitReader(compBytes);
    var out = [];
    var finalBlock = 0;

    var litTreeFixed = null, distTreeFixed = null;

    while (!finalBlock) {
      finalBlock = br.readBits(1);
      var btype = br.readBits(2);

      if (btype === 0) {
        br.alignByte();
        if (br.pos + 4 > br.bytes.length) throw new Error('Bad stored block');
        var len = br.bytes[br.pos] | (br.bytes[br.pos + 1] << 8);
        var nlen = br.bytes[br.pos + 2] | (br.bytes[br.pos + 3] << 8);
        br.pos += 4;
        if (((len ^ 0xFFFF) & 0xFFFF) !== (nlen & 0xFFFF)) throw new Error('Bad stored block len');
        for (var i = 0; i < len; i++) out.push(br.bytes[br.pos++]);
      } else {
        var litTree, distTree;

        if (btype === 1) {
          if (!litTreeFixed) litTreeFixed = fixedLitLenTree();
          if (!distTreeFixed) distTreeFixed = fixedDistTree();
          litTree = litTreeFixed;
          distTree = distTreeFixed;
        } else if (btype === 2) {
          var HLIT = br.readBits(5) + 257;
          var HDIST = br.readBits(5) + 1;
          var HCLEN = br.readBits(4) + 4;

          var order = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
          var clen = new Array(19);
          for (var z = 0; z < 19; z++) clen[z] = 0;
          for (z = 0; z < HCLEN; z++) clen[order[z]] = br.readBits(3);
          var clenTree = buildHuffman(clen);

          var all = new Array(HLIT + HDIST);
          var idx = 0;
          while (idx < all.length) {
            var sym = decodeSym(br, clenTree);
            if (sym <= 15) {
              all[idx++] = sym;
            } else if (sym === 16) {
              var repeat = br.readBits(2) + 3;
              var prev = idx ? all[idx - 1] : 0;
              while (repeat-- && idx < all.length) all[idx++] = prev;
            } else if (sym === 17) {
              repeat = br.readBits(3) + 3;
              while (repeat-- && idx < all.length) all[idx++] = 0;
            } else if (sym === 18) {
              repeat = br.readBits(7) + 11;
              while (repeat-- && idx < all.length) all[idx++] = 0;
            } else {
              throw new Error('Bad code length symbol');
            }
          }

          var litLens = all.slice(0, HLIT);
          var distLens = all.slice(HLIT);

          litTree = buildHuffman(litLens);
          distTree = buildHuffman(distLens);
        } else {
          throw new Error('Unsupported block type');
        }

        while (true) {
          var s = decodeSym(br, litTree);
          if (s < 256) {
            out.push(s);
          } else if (s === 256) {
            break;
          } else {
            var li = s - 257;
            var length = LEN_BASE[li] + br.readBits(LEN_EXTRA[li]);

            var ds = decodeSym(br, distTree);
            var dist = DIST_BASE[ds] + br.readBits(DIST_EXTRA[ds]);

            var start = out.length - dist;
            if (start < 0) throw new Error('Bad distance');

            for (var k = 0; k < length; k++) out.push(out[start + k]);
          }
        }
      }
    }

    var u8 = new Uint8Array(out.length);
    for (var j = 0; j < out.length; j++) u8[j] = out[j] & 0xFF;

    if (expectedSize && u8.length !== expectedSize) {
      // keep anyway
    }
    return u8;
  }

  // ----------------------------
  // XLSX (OOXML) minimal reader
  // ----------------------------
  function xmlUnescape(s) {
    return (s || '')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }

  function parseSharedStrings(xml) {
    var out = [];
    var reSi = /<si[\s\S]*?<\/si>/g;
    var m;
    while ((m = reSi.exec(xml)) !== null) {
      var si = m[0];
      var reT = /<t[^>]*>([\s\S]*?)<\/t>/g;
      var mt, s = '';
      while ((mt = reT.exec(si)) !== null) s += xmlUnescape(mt[1]);
      out.push(s);
    }
    return out;
  }

  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }

  function numToCol(n) {
    var s = '';
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function decodeA1(a1) {
    var m = /^([A-Z]+)(\d+)$/.exec(a1);
    if (!m) return { c: 0, r: 0 };
    return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
  }

  function parseDimension(xml) {
    var m = /<dimension[^>]*ref="([^"]+)"/.exec(xml);
    if (!m) return { s: { c: 1, r: 1 }, e: { c: 26, r: 50 } };
    var ref = m[1];
    var parts = ref.split(':');
    var a = decodeA1(parts[0]);
    var b = parts.length > 1 ? decodeA1(parts[1]) : a;
    return { s: { c: a.c, r: a.r }, e: { c: b.c, r: b.r } };
  }

  function parseSheet(xml, sharedStrings) {
    var cells = {};
    var reCell = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    var mc;
    while ((mc = reCell.exec(xml)) !== null) {
      var attrs = mc[1];
      var inner = mc[2];

      var mr = /r="([^"]+)"/.exec(attrs);
      if (!mr) continue;
      var addr = mr[1];

      var mt = /t="([^"]+)"/.exec(attrs);
      var t = mt ? mt[1] : '';

      var vMatch = /<v>([\s\S]*?)<\/v>/.exec(inner);
      var v = vMatch ? vMatch[1] : '';

      if (t === 's') {
        var idx = parseInt(v, 10);
        cells[addr] = (sharedStrings && sharedStrings[idx] !== undefined) ? sharedStrings[idx] : '';
      } else if (t === 'b') {
        cells[addr] = (v === '1') ? 'TRUE' : 'FALSE';
      } else if (t === 'str') {
        cells[addr] = xmlUnescape(v);
      } else if (t === 'inlineStr') {
        var it = /<is>[\s\S]*?<\/is>/.exec(inner);
        if (it) {
          var rt = /<t[^>]*>([\s\S]*?)<\/t>/.exec(it[0]);
          cells[addr] = rt ? xmlUnescape(rt[1]) : '';
        } else {
          cells[addr] = '';
        }
      } else {
        if (v === '') cells[addr] = '';
        else {
          var num = Number(v);
          cells[addr] = (isNaN(num) ? xmlUnescape(v) : num);
        }
      }
    }

    var hf = {};
    var mHF = /<headerFooter[\s\S]*?<\/headerFooter>/.exec(xml);
    if (mHF) {
      var block = mHF[0];
      var mOddH = /<oddHeader>([\s\S]*?)<\/oddHeader>/.exec(block);
      var mOddF = /<oddFooter>([\s\S]*?)<\/oddFooter>/.exec(block);
      hf.oddHeader = mOddH ? xmlUnescape(mOddH[1]) : '';
      hf.oddFooter = mOddF ? xmlUnescape(mOddF[1]) : '';
    }

    return { cells: cells, dim: parseDimension(xml), headerFooter: hf, rawXml: xml };
  }

  function openXlsx(arrayBuffer) {
    var bytes = new Uint8Array(arrayBuffer);
    var fileMap = unzipLocal(bytes);

    var wbBytes = fileMap['xl/workbook.xml'];
    if (!wbBytes || !wbBytes.length) throw new Error('workbook.xml saknas');
    var wbXml = decodeUtf8(wbBytes);

    var relsXml = '';
    if (fileMap['xl/_rels/workbook.xml.rels']) relsXml = decodeUtf8(fileMap['xl/_rels/workbook.xml.rels']);

    var sharedStrings = [];
    if (fileMap['xl/sharedStrings.xml']) sharedStrings = parseSharedStrings(decodeUtf8(fileMap['xl/sharedStrings.xml']));

    var ridToTarget = {};
    relsXml.replace(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/g, function (_, id, target) {
      ridToTarget[id] = target;
      return '';
    });

    var sheetDefs = [];
    wbXml.replace(/<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*\/?>/g, function (_, name, rid) {
      sheetDefs.push({ name: name, rid: rid });
      return '';
    });

    var sheets = [];
    for (var i = 0; i < sheetDefs.length; i++) {
      var target = ridToTarget[sheetDefs[i].rid] || '';
      if (!target) continue;
      target = target.replace(/^\//, '');
      var path = 'xl/' + target;
      var sBytes = fileMap[path];
      if (!sBytes) continue;
      var sXml = decodeUtf8(sBytes);
      var sheet = parseSheet(sXml, sharedStrings);
      sheet.name = sheetDefs[i].name;
      sheets.push(sheet);
    }

    return { sheets: sheets, fileMap: fileMap };
  }

  function getCell(sheet, addr) {
    if (!sheet || !sheet.cells) return '';
    var v = sheet.cells[addr];
    if (v === undefined || v === null) return '';
    return v;
  }

  function getText(sheet, addr) {
    var v = getCell(sheet, addr);
    if (typeof v === 'number') return String(v);
    return (v || '').toString();
  }

  function isDigitStart(s) {
    s = (s || '').toString().trim();
    if (!s) return false;
    var c = s.charAt(0);
    return c >= '0' && c <= '9';
  }

  function trim(s) { return (s || '').toString().replace(/^\s+|\s+$/g, ''); }

  function normalizeSig(s) {
    s = (s || '').toString().trim();
    if (!s) return '';
    s = s.replace(/\./g, '').replace(/\s+/g, ' ').toUpperCase();
    return s;
  }

  function keys(map) {
    var out = [];
    for (var k in map) if (map.hasOwnProperty(k)) out.push(k);
    out.sort();
    return out;
  }

  // ----------------------------
  // Seal Test analysis
  // ----------------------------
  function findHeaderCol(sheet, headerText, headerRow) {
    headerRow = headerRow || 1;
    var dim = sheet.dim || { s: { c: 1, r: 1 }, e: { c: 26, r: 50 } };
    var maxC = dim.e.c || 26;
    var needle = (headerText || '').toLowerCase();
    for (var c = 1; c <= maxC; c++) {
      var addr = numToCol(c) + headerRow;
      var v = (getText(sheet, addr) || '').toLowerCase();
      if (v === needle) return numToCol(c);
    }
    return '';
  }

  function parseCellAddr(addr) {
    var m = /^([A-Z]+)(\d+)$/i.exec(String(addr || '').trim());
    if (!m) return null;
    return { c: colToNum(m[1].toUpperCase()), r: parseInt(m[2], 10) };
  }

  function makeAddr(c, r) {
    if (!c || !r) return '';
    return numToCol(c) + String(r);
  }

  function addrRight(addr, dx) {
    var p = parseCellAddr(addr);
    if (!p) return '';
    return makeAddr(p.c + (dx || 1), p.r);
  }

  function addrDown(addr, dy) {
    var p = parseCellAddr(addr);
    if (!p) return '';
    return makeAddr(p.c, p.r + (dy || 1));
  }

  function findLabelAddr(sheet, re) {
    if (!sheet || !sheet.cells) return '';
    for (var k in sheet.cells) {
      if (!sheet.cells.hasOwnProperty(k)) continue;
      var txt = trim(getText(sheet, k));
      if (txt && re.test(txt)) return k;
    }
    return '';
  }

  function valueNear(sheet, labelRe) {
    var a = findLabelAddr(sheet, labelRe);
    if (!a) return '';

    // Prefer right cell (common for key/value on same row)
    var v = trim(getText(sheet, addrRight(a, 1)));
    if (v && !labelRe.test(v)) return v;

    // Then below (common for key on row, value on next row)
    v = trim(getText(sheet, addrDown(a, 1)));
    if (v && !labelRe.test(v)) return v;

    // Fallback: a bit further away
    v = trim(getText(sheet, addrRight(a, 2)));
    if (v && !labelRe.test(v)) return v;

    v = trim(getText(sheet, addrDown(a, 2)));
    if (v && !labelRe.test(v)) return v;

    return '';
  }

  function readSealHeader(sheet) {
    return {
      ROBAL: valueNear(sheet, /^ROBAL\b/i),
      PartNumber: valueNear(sheet, /^Part\s*Number\b/i),
      BatchNumber: valueNear(sheet, /^Batch\s*Number/i),
      CartridgeLsp: valueNear(sheet, /(Cartridge.*\(LSP\)|Cartridge\s*LSP|Cartridge\s*No\.?\s*\(LSP\)|Cartridge\s*Number)/i),
      PO: valueNear(sheet, /^PO\s*Number\b/i),
      AssayFamily: valueNear(sheet, /^Assay\s*Family\b/i),
      WeightLossSpec: valueNear(sheet, /Weight\s*Loss\s*Spec/i)
    };
  }

  function readSealPeople(sheet) {
    var testers = [];
    var sigs = [];

    var t = trim(valueNear(sheet, /Name\s+of\s+Tester/i));
    if (!t) t = trim(valueNear(sheet, /Inspected\s+by/i));
    if (t) testers.push(t);

    var s = trim(valueNear(sheet, /Print\s+Full\s+Name.*Sign.*Date/i));
    if (s) sigs.push(s);

    return { testers: testers, sigs: sigs };
  }

  function analyzeSealTest(xlsxObj, label) {
    var dataSheets = [];
    for (var i = 0; i < xlsxObj.sheets.length; i++) {
      var sh = xlsxObj.sheets[i];
      if (sh.name === 'Worksheet Instructions') continue;
      if (isDigitStart(getText(sh, 'H3'))) dataSheets.push(sh);
    }

    var header = null;
    if (dataSheets.length) header = readSealHeader(dataSheets[0]);

    var testers = {};
    var signatures = {};
    for (i = 0; i < dataSheets.length; i++) {
      var t = trim(valueNear(dataSheets[i], /Name\s+of\s+Tester/i) || valueNear(dataSheets[i], /Inspected\s+by/i) || getText(dataSheets[i], 'B43'));
      if (t) {
        var parts = t.split(',');
        for (var p = 0; p < parts.length; p++) {
          var s = trim(parts[p]);
          if (s) testers[s] = true;
        }
      }
      var sig = normalizeSig(getText(dataSheets[i], 'B47'));
      if (sig) signatures[sig] = true;
    }

    var viol = [];
    for (i = 0; i < dataSheets.length; i++) {
      var sheet = dataSheets[i];
      var cart = getText(sheet, 'H3');
      var obsCol = findHeaderCol(sheet, 'Observation', 1);
      if (!obsCol) obsCol = 'M';
      for (var r = 3; r <= 45; r++) {
        var avg = getCell(sheet, 'K' + r);
        var status = (getText(sheet, 'L' + r) || '').toUpperCase();
        var avgNum = (typeof avg === 'number') ? avg : Number(avg);
        var bad = false;
        var reason = '';
        if (!isNaN(avgNum) && avgNum <= -3.0) { bad = true; reason = 'MinusValue<=-3'; }
        if (status === 'FAIL') { bad = true; reason = 'FAIL'; }
        if (!bad) continue;

        viol.push({
          source: label,
          sheet: sheet.name,
          cartridge: cart,
          avg: (!isNaN(avgNum) ? avgNum : getText(sheet, 'K' + r)),
          status: status || '',
          reason: reason,
          initial: getText(sheet, 'H' + r),
          final: getText(sheet, 'I' + r),
          observation: getText(sheet, obsCol + r)
        });
      }
    }

    return {
      label: label,
      sheetCount: xlsxObj.sheets.length,
      dataSheetCount: dataSheets.length,
      header: header,
      testers: keys(testers),
      signatures: keys(signatures),
      violations: viol
    };
  }

  // ----------------------------
  // Worksheet analysis (minimal)
  // ----------------------------
  function analyzeWorksheet(xlsxObj) {
    var sheet = null;
    for (var i = 0; i < xlsxObj.sheets.length; i++) {
      if (xlsxObj.sheets[i].name === 'Test Summary') { sheet = xlsxObj.sheets[i]; break; }
    }
    if (!sheet && xlsxObj.sheets.length) sheet = xlsxObj.sheets[0];
    if (!sheet) return { ok: false, error: 'Worksheet: inga sheets hittades' };

    return {
      ok: true,
      sheetName: sheet.name,
      partNumber: getText(sheet, 'B3'),
      cartridgeLsp: getText(sheet, 'B4')
    };
  }

  // ----------------------------
  // Render results
  // ----------------------------
  function escapeHtml(s) {
    s = (s === undefined || s === null) ? '' : String(s);
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  }

  function renderDupList(title, arr) {
    if (!arr || !arr.length) return el('div', { 'class': 'small muted', text: title + ': inga dubletter.' });
    var wrap = el('div', { 'class': 'small' });
    wrap.appendChild(el('div', { 'class': 'muted', text: title + ': ' + arr.length + ' st.' }));
    var t = el('table');
    var th = el('thead');
    var tr = el('tr');
    tr.appendChild(el('th', { text: 'Value' }));
    tr.appendChild(el('th', { text: 'Count' }));
    th.appendChild(tr);
    t.appendChild(th);

    var tb = el('tbody');
    for (var i = 0; i < Math.min(arr.length, 20); i++) {
      var r = el('tr');
      r.appendChild(el('td', { 'class': 'mono', text: arr[i].key }));
      r.appendChild(el('td', { 'class': 'mono', text: String(arr[i].count) }));
      tb.appendChild(r);
    }
    t.appendChild(tb);
    wrap.appendChild(t);
    return wrap;
  }

    function renderResults(model) {
    var out = $('results');
    clearNode(out);

    out.appendChild(el('h2', { text: 'Resultat' }));

    if (!model || !model.csv || !model.csv.ok) {
      out.appendChild(el('div', { 'class': 'small err', text: 'Ingen CSV-data att visa.' }));
      return;
    }

    var csv = model.csv;

    // Overall status
    var hasDupSample = (csv.duplicateSamples && csv.duplicateSamples.length);
    var hasDupCart = (csv.duplicateCarts && csv.duplicateCarts.length);
    var hasInvalid = (csv.invalid && csv.invalid.length);
    var hasCtrlDev = (csv.ctrl && csv.ctrl.deviations && csv.ctrl.deviations.length);
    var hasMissingBags = (csv.bag && csv.bag.missing && csv.bag.missing.length);

    var top = el('div', { 'class': 'card' });
    var pills = el('div');
    pills.appendChild(pill('Rows: ' + csv.rowCount, 'ok'));
    if (csv.assay) pills.appendChild(pill('Assay: ' + csv.assay, 'ok'));
    if (csv.assayVer) pills.appendChild(pill('Ver: ' + csv.assayVer, 'ok'));
    if (csv.lot) pills.appendChild(pill('Lot: ' + csv.lot, 'ok'));

    if (hasCtrlDev) pills.appendChild(pill('Control deviations: ' + csv.ctrl.deviations.length, 'bad'));
    else pills.appendChild(pill('Control deviations: 0', 'ok'));

    if (hasInvalid) pills.appendChild(pill('Invalid/NoResult/Error: ' + csv.invalid.length, 'warn'));
    else pills.appendChild(pill('Invalid/NoResult/Error: 0', 'ok'));

    if (hasDupSample) pills.appendChild(pill('Dup SampleID: ' + csv.duplicateSamples.length, 'warn'));
    else pills.appendChild(pill('Dup SampleID: 0', 'ok'));

    if (hasDupCart) pills.appendChild(pill('Dup Cartridge S/N: ' + csv.duplicateCarts.length, 'warn'));
    else pills.appendChild(pill('Dup Cartridge S/N: 0', 'ok'));

    if (hasMissingBags) pills.appendChild(pill('Missing bags: ' + csv.bag.missing.length, 'warn'));
    else pills.appendChild(pill('Missing bags: 0', 'ok'));

    top.appendChild(pills);
    top.appendChild(el('div', { 'class': 'small muted', text: 'Delimiter: ' + csv.delim + ' | Header rad: ' + (csv.headerIdx + 1) }));
    out.appendChild(top);

    if (model._xlsxSelected) {
      out.appendChild(el('div', { 'class': 'small muted', text: 'Obs: XLSX-filer är valda, men v7 kör endast CSV-validering (XLSX kommer i nästa steg).' }));
    }

    // Helper: render table
    function renderTable(columns, rows, rowToCells) {
      var tbl = el('table');
      var thead = el('thead');
      var trh = el('tr');
      for (var c = 0; c < columns.length; c++) trh.appendChild(el('th', { text: columns[c] }));
      thead.appendChild(trh);
      tbl.appendChild(thead);

      var tbody = el('tbody');
      for (var r = 0; r < rows.length; r++) {
        var tr = el('tr');
        var cells = rowToCells(rows[r]);
        for (var k = 0; k < cells.length; k++) tr.appendChild(el('td', { text: cells[k] }));
        tbody.appendChild(tr);
      }
      tbl.appendChild(tbody);
      return tbl;
    }

    // Section: duplicates
    var secDup = el('div', { 'class': 'card' });
    secDup.appendChild(el('h3', { text: 'Dubletter' }));

    if (!hasDupSample && !hasDupCart) {
      secDup.appendChild(el('div', { 'class': 'small muted', text: 'Inga dubletter hittades.' }));
    } else {
      if (hasDupSample) {
        secDup.appendChild(el('div', { 'class': 'small', text: 'Sample ID (>' + 1 + '):' }));
        secDup.appendChild(renderTable(['Sample ID', 'Count'], csv.duplicateSamples, function (x) { return [x.value, String(x.count)]; }));
      }
      if (hasDupCart) {
        secDup.appendChild(el('div', { 'class': 'small', text: 'Cartridge S/N (>' + 1 + '):' }));
        secDup.appendChild(renderTable(['Cartridge S/N', 'Count'], csv.duplicateCarts, function (x) { return [x.value, String(x.count)]; }));
      }
    }
    out.appendChild(secDup);

    // Section: invalid results
    var secInv = el('div', { 'class': 'card' });
    secInv.appendChild(el('h3', { text: 'Invalid / No Result / Error' }));
    if (!hasInvalid) {
      secInv.appendChild(el('div', { 'class': 'small muted', text: 'Inga rader med NO RESULT / INVALID / ERROR hittades.' }));
    } else {
      secInv.appendChild(renderTable(
        ['Sample ID', 'Ctrl', 'Bag', 'Cartridge S/N', 'Test Type', 'Test Result'],
        csv.invalid,
        function (r) { return [r.sampleId, r.ctrl, (r.bag === null ? '' : String(r.bag)), r.cartSn, r.testType, r.testResult]; }
      ));
    }
    out.appendChild(secInv);

    // Section: control expectations + deviations
    var secCtrl = el('div', { 'class': 'card' });
    secCtrl.appendChild(el('h3', { text: 'Kontroll-grupper (majoritet) + avvikelser' }));

    // Summary per ctrl
    var ctrlKeys = [];
    for (var ck in csv.ctrl.counts) if (csv.ctrl.counts.hasOwnProperty(ck)) ctrlKeys.push(ck);
    ctrlKeys.sort(function (a, b) {
      var ai = parseInt(a, 10), bi = parseInt(b, 10);
      if (!isNaN(ai) && !isNaN(bi)) return ai - bi;
      return String(a).localeCompare(String(b));
    });

    var ctrlRows = [];
    for (var ci = 0; ci < ctrlKeys.length; ci++) {
      var key = ctrlKeys[ci];
      var c = csv.ctrl.counts[key];
      var exp = csv.ctrl.expected[key] || 'unknown';
      // prefix summary (top 2)
      var pmap = csv.ctrl.prefixes[key] || {};
      var plist = [];
      for (var pk in pmap) if (pmap.hasOwnProperty(pk)) plist.push({ k: pk, n: pmap[pk] });
      plist.sort(function (a, b) { return b.n - a.n; });
      var ptxt = '';
      for (var pi = 0; pi < plist.length && pi < 2; pi++) {
        ptxt += (ptxt ? ', ' : '') + plist[pi].k + ' (' + plist[pi].n + ')';
      }

      ctrlRows.push({
        ctrl: key,
        expected: exp,
        detected: c.detected || 0,
        notDetected: c.not_detected || 0,
        unknown: c.unknown || 0,
        total: c.total || 0,
        prefixes: ptxt
      });
    }

    secCtrl.appendChild(renderTable(
      ['Ctrl', 'Expected', 'Detected', 'Not detected', 'Unknown', 'Total', 'Top prefixes'],
      ctrlRows,
      function (r) { return [r.ctrl, r.expected, String(r.detected), String(r.notDetected), String(r.unknown), String(r.total), r.prefixes]; }
    ));

    if (!hasCtrlDev) {
      secCtrl.appendChild(el('div', { 'class': 'small muted', text: 'Inga avvikelser mot majoritetsförväntan i kontroll-grupperna.' }));
    } else {
      secCtrl.appendChild(el('div', { 'class': 'small', text: 'Avvikelser (Expected != Observed):' }));
      secCtrl.appendChild(renderTable(
        ['Ctrl', 'Expected', 'Observed', 'Sample ID', 'Bag', 'Cartridge S/N', 'Test Type', 'Test Result'],
        csv.ctrl.deviations,
        function (r) {
          return [
            r.ctrl,
            r.expected,
            r.detectedClass,
            r.sampleId,
            (r.bag === null ? '' : String(r.bag)),
            r.cartSn,
            r.testType,
            r.testResult
          ];
        }
      ));
    }

    out.appendChild(secCtrl);

    // Section: bag coverage
    var secBag = el('div', { 'class': 'card' });
    secBag.appendChild(el('h3', { text: 'Bag coverage' }));
    if (csv.bag.min === null) {
      secBag.appendChild(el('div', { 'class': 'small muted', text: 'Kunde inte tolka bag-nummer från Sample ID.' }));
    } else {
      secBag.appendChild(el('div', { 'class': 'small', text: 'Bag range: ' + csv.bag.min + ' → ' + csv.bag.max }));
      if (hasMissingBags) {
        secBag.appendChild(el('div', { 'class': 'small err', text: 'Missing bags: ' + csv.bag.missing.join(', ') }));
      } else {
        secBag.appendChild(el('div', { 'class': 'small muted', text: 'Inga saknade bag-nummer inom spannet.' }));
      }
    }
    out.appendChild(secBag);

    // Debug footer
    out.appendChild(el('div', { 'class': 'small muted', text: 'Tips: Om något ser fel ut, öppna DevTools (F12) och kolla Console.' }));
  }


  // ----------------------------
  // Read files + run analysis
  // ----------------------------
  function readFileAsArrayBuffer(file, cb) {
    var fr = new FileReader();
    fr.onload = function () { cb(null, fr.result); };
    fr.onerror = function () { cb(new Error('FileReader error: ' + file.name)); };
    fr.readAsArrayBuffer(file);
  }

  function readFileAsText(file, cb) {
    var fr = new FileReader();

    // Backwards compatible: if cb provided, use callback style.
    if (typeof cb === 'function') {
      fr.onload = function () { cb(null, fr.result); };
      fr.onerror = function () { cb(new Error('FileReader error: ' + ((file && file.name) ? file.name : 'file'))); };
      fr.readAsText(file);
      return;
    }

    // Promise style (preferred)
    if (typeof Promise === 'undefined') {
      throw new Error('Promise is not supported in this browser.');
    }

    return new Promise(function (resolve, reject) {
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(new Error('FileReader error: ' + ((file && file.name) ? file.name : 'file'))); };
      fr.readAsText(file);
    });
  }

  function pickByType(typeKey) {
    for (var i = 0; i < files.length; i++) {
      var tk = files[i].typeKey;
      if (tk === 'auto') tk = files[i].detectedKey;
      if (tk === typeKey) return files[i].file;
    }
    return null;
  }

    function runAnalysis() {
    var btn = $('runBtn');
    btn.disabled = true;
    setStatus('Analyserar...', false);

    var fCsv = pickByType('testSummary');
    var fNeg = pickByType('sealNeg');
    var fPos = pickByType('sealPos');
    var fWs = pickByType('worksheet');

    if (!fCsv) {
      setStatus('Saknar Test Summary (csv).', true);
      validateReady();
      return;
    }

    clearNode($('results'));
    $('results').appendChild(el('div', { 'class': 'small muted', text: 'Läser och analyserar CSV...' }));

    readFileAsText(fCsv).then(function (csvText) {
      var csv = parseCsv(csvText);
      if (!csv.ok) {
        throw new Error(csv.error || 'CSV-parse misslyckades.');
      }

      // v7 milestone: CSV validation first.
      // XLSX parsing will be re-introduced once we have a reliable offline XLSX reader.
      var model = { csv: csv, sealNeg: null, sealPos: null, worksheet: null };

      // If the user selected XLSX files, show a friendly note (don’t try to parse yet).
      model._xlsxSelected = !!(fNeg || fPos || fWs);

      renderResults(model);
      setStatus('OK – CSV-analys klar.', false);
      validateReady();
    }).catch(function (e) {
      var msg = (e && e.message) ? e.message : String(e);
      setStatus('Fel: ' + msg, true);
      clearNode($('results'));
      $('results').appendChild(el('div', { 'class': 'small err', text: msg }));
      validateReady();
    });
  }


  // ----------------------------
  // Wire UI
  // ----------------------------
  function init() {
    $('pickBtn').addEventListener('click', function () { $('fileInput').click(); });
    $('fileInput').addEventListener('change', function (ev) {
      if (ev.target && ev.target.files) addFiles(ev.target.files);
      $('fileInput').value = '';
    });

    $('clearBtn').addEventListener('click', function () {
      files = [];
      refreshFileList();
      $('results').textContent = 'Ingen analys kord an.';
      setStatus('', false);
    });

    $('runBtn').addEventListener('click', runAnalysis);

    var drop = $('drop');
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.style.borderColor = 'rgba(31,111,235,0.9)'; });
    drop.addEventListener('dragleave', function () { drop.style.borderColor = 'rgba(255,255,255,0.16)'; });
    drop.addEventListener('drop', function (e) {
      e.preventDefault();
      drop.style.borderColor = 'rgba(255,255,255,0.16)';
      if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
    });

    refreshFileList();
  }

  init();

})();