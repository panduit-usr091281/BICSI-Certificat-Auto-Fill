/* ================================================================
   cert_generator_app.js
   BICSI Certificate Auto-Fill — client-side certificate generation.
   All processing runs in the browser; nothing is uploaded anywhere.
   ================================================================ */

(function () {
  'use strict';

  /* -------- state -------- */
  var worksheetData = null;    // { columns: string[], rows: object[] }
  var templateBuffer = null;   // ArrayBuffer of the uploaded .docx

  /* -------- DOM refs -------- */
  var xlsxInput        = document.getElementById('xlsx-input');
  var docxInput        = document.getElementById('docx-input');
  var xlsxZone         = document.getElementById('xlsx-zone');
  var docxZone         = document.getElementById('docx-zone');
  var xlsxLabel        = document.getElementById('xlsx-label');
  var docxLabel        = document.getElementById('docx-label');

  var mappingSection   = document.getElementById('step-mapping');
  var dataPreview      = document.getElementById('data-preview');
  var nameCol          = document.getElementById('name-col');
  var orgCol           = document.getElementById('org-col');
  var dateCol          = document.getElementById('date-col');
  var namePH           = document.getElementById('name-ph');
  var orgPH            = document.getElementById('org-ph');
  var datePH           = document.getElementById('date-ph');

  var generateSection  = document.getElementById('step-generate');
  var outputFormat     = document.getElementById('output-format');
  var formatNote       = document.getElementById('format-note');
  var btnGenerate      = document.getElementById('btn-generate');
  var progressArea     = document.getElementById('progress-area');
  var progressFill     = document.getElementById('progress-fill');
  var statusText       = document.getElementById('status-text');

  var downloadSection  = document.getElementById('step-download');
  var resultSummary    = document.getElementById('result-summary');
  var btnDownload      = document.getElementById('btn-download');

  var renderContainer  = document.getElementById('render-container');

  /* ==================================================================
     1. FILE UPLOADS
     ================================================================== */

  xlsxInput.addEventListener('change', function (e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var data = new Uint8Array(ev.target.result);
        var wb   = XLSX.read(data, { type: 'array', cellDates: true });
        var sheet = wb.Sheets[wb.SheetNames[0]];
        var json  = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (json.length === 0) { alert('The spreadsheet appears to be empty.'); return; }

        worksheetData = { columns: Object.keys(json[0]), rows: json };
        xlsxLabel.textContent = file.name + ' (' + json.length + ' rows)';
        xlsxZone.classList.add('uploaded');
        tryShowMapping();
      } catch (err) {
        alert('Could not read spreadsheet: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  docxInput.addEventListener('change', function (e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function (ev) {
      templateBuffer = ev.target.result;
      docxLabel.textContent = file.name;
      docxZone.classList.add('uploaded');
      tryShowMapping();
    };
    reader.readAsArrayBuffer(file);
  });

  /* ==================================================================
     2. COLUMN MAPPING
     ================================================================== */

  function tryShowMapping() {
    if (!worksheetData || !templateBuffer) return;
    show(mappingSection);
    show(generateSection);
    buildPreviewTable();
    populateSelects();
  }

  function buildPreviewTable() {
    var cols = worksheetData.columns;
    var rows = worksheetData.rows.slice(0, 5);
    var html = '<thead><tr>';
    cols.forEach(function (c) { html += '<th>' + esc(c) + '</th>'; });
    html += '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr>';
      cols.forEach(function (c) { html += '<td>' + esc(displayValue(r[c])) + '</td>'; });
      html += '</tr>';
    });
    html += '</tbody>';
    dataPreview.innerHTML = html;
  }

  function populateSelects() {
    var cols = worksheetData.columns;
    [nameCol, orgCol, dateCol].forEach(function (sel) {
      sel.innerHTML = '';
      cols.forEach(function (c) {
        var opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        sel.appendChild(opt);
      });
    });
    autoDetect(nameCol, cols, ['name', 'full name', 'student', 'participant', 'attendee']);
    autoDetect(orgCol,  cols, ['organization', 'org', 'company', 'employer']);
    autoDetect(dateCol, cols, ['date', 'attendance', 'date of attendance', 'completion']);
  }

  function autoDetect(selectEl, cols, keywords) {
    for (var i = 0; i < cols.length; i++) {
      var low = cols[i].toLowerCase();
      for (var k = 0; k < keywords.length; k++) {
        if (low.indexOf(keywords[k]) !== -1) { selectEl.value = cols[i]; return; }
      }
    }
  }

  /* Show / hide the PDF format note based on selection */
  outputFormat.addEventListener('change', function () {
    formatNote.style.display = (outputFormat.value === 'docx') ? 'none' : '';
  });

  /* ==================================================================
     3. DOCX TEMPLATE FILLING (pure string replacement — no XML parsing)
     Avoids DOMParser/XMLSerializer which can destroy text boxes, lines,
     drawings, and other rich OOXML features.
     ================================================================== */

  /**
   * Fill the template with one set of replacement values.
   * Returns a Promise<ArrayBuffer> of the completed .docx.
   */
  async function fillTemplate(templateBuf, replacements) {
    var zip = await JSZip.loadAsync(templateBuf);

    // Collect XML parts that can contain text
    var parts = ['word/document.xml'];
    zip.forEach(function (path) {
      if (/^word\/(header|footer)\d*\.xml$/.test(path)) parts.push(path);
    });

    // Build safe replacements (XML-escape the values)
    var safeReplacements = {};
    for (var key in replacements) {
      safeReplacements[key] = escapeXml(replacements[key]);
    }

    for (var i = 0; i < parts.length; i++) {
      var f = zip.file(parts[i]);
      if (!f) continue;
      var xml = await f.async('string');
      xml = replacePlaceholders(xml, safeReplacements);
      zip.file(parts[i], xml);
    }

    return zip.generateAsync({ type: 'arraybuffer' });
  }

  /** Escape special XML characters so replacement values don't break the doc. */
  function escapeXml(str) {
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');
  }

  /**
   * Replace placeholders in the raw XML string without ever parsing it.
   * This preserves every byte of the original structure (text boxes, VML
   * shapes, drawings, SmartArt, etc.) — only the placeholder text changes.
   *
   * Strategy:
   *  1. Direct string replacement (covers placeholders inside a single <w:t>).
   *  2. Cross-run handler: if a placeholder is split across multiple <w:t>
   *     elements by Word, we find it in each <w:p> block and stitch the
   *     runs back together.
   */
  function replacePlaceholders(xmlStr, safeReplacements) {
    for (var ph in safeReplacements) {
      var val = safeReplacements[ph];

      // --- Pass 1: simple same-run replacement ---
      if (xmlStr.indexOf(ph) !== -1) {
        xmlStr = xmlStr.split(ph).join(val);
      }

      // --- Pass 2: cross-run replacement ---
      // Only needed if any <w:p> block's concatenated <w:t> text still
      // contains the placeholder after pass 1 (i.e. it was split across runs).
      xmlStr = replaceCrossRun(xmlStr, ph, val);
    }
    return xmlStr;
  }

  /**
   * Handle the case where Word split a placeholder across multiple <w:r>/<w:t>
   * elements.  Works entirely on the raw XML string — no DOM parsing.
   *
   * For each <w:p>…</w:p> block:
   *   1. Extract all <w:t …>text</w:t> fragments and their string offsets.
   *   2. Concatenate the text.  If the placeholder isn't present, skip.
   *   3. Put the replacement value into the first <w:t> that touched the
   *      placeholder and empty the remaining fragments.
   */
  function replaceCrossRun(xmlStr, placeholder, value) {
    // Regex to match paragraph blocks (non-greedy within <w:p>…</w:p>)
    var paraRe = /<w:p[\s>][\s\S]*?<\/w:p>/g;

    return xmlStr.replace(paraRe, function (paraXml) {
      // Collect every <w:t …>text</w:t> inside this paragraph
      var tRe    = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g;
      var pieces = [];   // { open, text, close, index, length } in order
      var m;

      while ((m = tRe.exec(paraXml)) !== null) {
        pieces.push({
          open:   m[1],          // e.g.  <w:t xml:space="preserve">
          text:   m[2],          // the actual text content
          close:  m[3],          // </w:t>
          start:  m.index,       // offset inside paraXml
          len:    m[0].length    // full match length
        });
      }

      if (pieces.length === 0) return paraXml;

      // Build concatenated plain text
      var fullText = pieces.map(function (p) { return p.text; }).join('');
      var idx = fullText.indexOf(placeholder);
      if (idx === -1) return paraXml;  // nothing to do

      // Walk through occurrences
      while (idx !== -1) {
        var endIdx = idx + placeholder.length;

        // Map character positions → piece indexes
        var charPos = 0;
        var firstPiece = -1, firstOff = 0, lastPiece = -1, lastOff = 0;

        for (var pi = 0; pi < pieces.length; pi++) {
          var pLen = pieces[pi].text.length;
          if (firstPiece === -1 && charPos + pLen > idx) {
            firstPiece = pi;
            firstOff   = idx - charPos;
          }
          if (charPos + pLen >= endIdx) {
            lastPiece = pi;
            lastOff   = endIdx - charPos;
            break;
          }
          charPos += pLen;
        }

        // Rewrite the text content of the affected pieces
        var before = pieces[firstPiece].text.substring(0, firstOff);
        var after  = pieces[lastPiece].text.substring(lastOff);

        pieces[firstPiece].text = before + value + (firstPiece === lastPiece ? after : '');

        if (firstPiece !== lastPiece) {
          for (var ci = firstPiece + 1; ci < lastPiece; ci++) {
            pieces[ci].text = '';
          }
          pieces[lastPiece].text = after;
        }

        // Recheck for another occurrence
        fullText = pieces.map(function (p) { return p.text; }).join('');
        idx = fullText.indexOf(placeholder);
      }

      // Rebuild the paragraph XML from right to left (so offsets stay valid)
      var result = paraXml;
      for (var ri = pieces.length - 1; ri >= 0; ri--) {
        var pc    = pieces[ri];
        var rebuilt = pc.open + pc.text + pc.close;
        result = result.substring(0, pc.start) + rebuilt + result.substring(pc.start + pc.len);
      }

      return result;
    });
  }

  /* ==================================================================
     4. PDF GENERATION (docx-preview → html2canvas → jsPDF)
     ================================================================== */

  /**
   * Convert a filled DOCX ArrayBuffer into a PDF ArrayBuffer.
   * Renders the DOCX to an off-screen container, captures it as an image,
   * then wraps it in a single-page PDF.
   */
  async function docxToPdf(docxBuf) {
    renderContainer.innerHTML = '';

    await docx.renderAsync(docxBuf, renderContainer, null, {
      className: 'cert-render',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      renderHeaders: true,
      renderFooters: true,
      breakPages: false
    });

    // Give images / fonts a moment to paint
    await delay(400);

    var target = renderContainer.querySelector('.docx-wrapper') || renderContainer;

    var canvas = await html2canvas(target, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      logging: false,
      backgroundColor: '#ffffff'
    });

    var imgData  = canvas.toDataURL('image/jpeg', 0.95);
    var pxW      = canvas.width;
    var pxH      = canvas.height;
    var orient   = pxW > pxH ? 'landscape' : 'portrait';

    var jsPDFLib = window.jspdf.jsPDF;
    var pdf = new jsPDFLib({
      orientation: orient,
      unit: 'px',
      format: [pxW / 2, pxH / 2],
      hotfixes: ['px_scaling']
    });
    pdf.addImage(imgData, 'JPEG', 0, 0, pxW / 2, pxH / 2);

    renderContainer.innerHTML = '';
    return pdf.output('arraybuffer');
  }

  /* ==================================================================
     5. MAIN GENERATION FLOW
     ================================================================== */

  btnGenerate.addEventListener('click', generateCertificates);

  async function generateCertificates() {
    var fmt      = outputFormat.value;          // 'pdf' | 'docx' | 'both'
    var nCol     = nameCol.value;
    var oCol     = orgCol.value;
    var dCol     = dateCol.value;
    var nPH      = namePH.value;
    var oPH      = orgPH.value;
    var dPH      = datePH.value;

    if (!nPH || !oPH || !dPH) {
      alert('Please fill in all three placeholder fields.');
      return;
    }

    // Check that docx-preview is available when PDF is requested
    if (fmt !== 'docx' && typeof docx === 'undefined') {
      alert('The PDF rendering library (docx-preview) failed to load.\n' +
            'Please check your internet connection or choose DOCX format.');
      return;
    }

    btnGenerate.disabled = true;
    show(progressArea);
    hide(downloadSection);

    var outZip     = new JSZip();
    var generated  = 0;
    var skipped    = 0;
    var total      = worksheetData.rows.length;
    var usedNames  = {};

    try {
      for (var i = 0; i < total; i++) {
        var row  = worksheetData.rows[i];
        var name = String(row[nCol] != null ? row[nCol] : '').trim();
        var org  = String(row[oCol] != null ? row[oCol] : '').trim();
        var dateVal = row[dCol];

        if (!name) { skipped++; setProgress(i + 1, total, 'Skipping empty row…'); continue; }

        var dateStr = formatDateValue(dateVal);
        var repl    = {};
        repl[nPH]   = name;
        repl[oPH]   = org;
        repl[dPH]   = dateStr;

        var safeName = sanitize(name);
        var safeOrg  = sanitize(org).substring(0, 20);
        var baseName = 'Certificate_' + safeName + '_' + safeOrg;

        // Ensure unique filename
        if (usedNames[baseName]) {
          usedNames[baseName]++;
          baseName += '_' + usedNames[baseName];
        } else {
          usedNames[baseName] = 1;
        }

        setProgress(i + 1, total, 'Filling template for ' + name + '…');
        var filledDocx = await fillTemplate(templateBuffer, repl);

        if (fmt === 'docx' || fmt === 'both') {
          outZip.file(baseName + '.docx', filledDocx);
        }

        if (fmt === 'pdf' || fmt === 'both') {
          setProgress(i + 1, total, 'Rendering PDF for ' + name + '…');
          var pdfBuf = await docxToPdf(filledDocx);
          outZip.file(baseName + '.pdf', pdfBuf);
        }

        generated++;
      }

      setProgress(total, total, 'Packaging certificates…');
      var blob = await outZip.generateAsync({ type: 'blob' });

      resultSummary.textContent =
        generated + ' certificate(s) generated' +
        (skipped ? ', ' + skipped + ' row(s) skipped.' : '.');
      show(downloadSection);

      btnDownload.onclick = function () { saveAs(blob, 'BICSI_Certificates.zip'); };

      statusText.textContent = 'Done!';
    } catch (err) {
      alert('Error during generation: ' + err.message);
      console.error(err);
      statusText.textContent = 'Generation failed.';
    }

    btnGenerate.disabled = false;
  }

  /* ==================================================================
     6. UTILITIES
     ================================================================== */

  function sanitize(text) {
    return text.replace(/[<>:"/\\|?*\x00-\x1f]/g, '').replace(/\s+/g, ' ').trim() || 'Unknown';
  }

  function formatDateValue(val) {
    if (val == null || val === '') return '';

    if (val instanceof Date) {
      if (isNaN(val.getTime())) return String(val);
      return monthName(val.getUTCMonth()) + ' ' + val.getUTCDate() + ', ' + val.getUTCFullYear();
    }

    if (typeof val === 'number') {
      // Excel serial date
      try {
        var p = XLSX.SSF.parse_date_code(val);
        return monthName(p.m - 1) + ' ' + p.d + ', ' + p.y;
      } catch (_) { return String(val); }
    }

    // String — try parsing, else return as-is
    if (typeof val === 'string') {
      var d = new Date(val);
      if (!isNaN(d.getTime())) {
        return monthName(d.getMonth()) + ' ' + d.getDate() + ', ' + d.getFullYear();
      }
      return val;
    }

    return String(val);
  }

  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  function monthName(i) { return MONTHS[i]; }

  function displayValue(v) {
    if (v instanceof Date) return formatDateValue(v);
    if (v == null) return '';
    return String(v);
  }

  function esc(s) {
    var d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
  }

  function show(el) { el.classList.remove('hidden'); }
  function hide(el) { el.classList.add('hidden'); }

  function setProgress(current, total, msg) {
    var pct = Math.round((current / total) * 100);
    progressFill.style.width = pct + '%';
    statusText.textContent = msg;
  }

  function delay(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

})();
