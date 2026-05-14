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
     3. DOCX TEMPLATE FILLING (XML-level text replacement)
     ================================================================== */

  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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

    for (var i = 0; i < parts.length; i++) {
      var f = zip.file(parts[i]);
      if (!f) continue;
      var xml = await f.async('string');
      xml = replacePlaceholders(xml, replacements);
      zip.file(parts[i], xml);
    }

    return zip.generateAsync({ type: 'arraybuffer' });
  }

  /** Parse an XML string, replace placeholders in every <w:p>, serialize back. */
  function replacePlaceholders(xmlStr, replacements) {
    var parser = new DOMParser();
    var doc    = parser.parseFromString(xmlStr, 'application/xml');
    var paras  = doc.getElementsByTagNameNS(W_NS, 'p');

    for (var i = 0; i < paras.length; i++) {
      replaceInParagraph(paras[i], replacements);
    }

    return new XMLSerializer().serializeToString(doc);
  }

  /**
   * Replace placeholders inside a single <w:p>, handling the case where
   * Word splits a placeholder across multiple <w:r>/<w:t> elements.
   * Preserves all run-level formatting (font, size, bold, color, etc.).
   */
  function replaceInParagraph(para, replacements) {
    var runs = para.getElementsByTagNameNS(W_NS, 'r');
    if (runs.length === 0) return;

    // Collect every <w:t> and its parent run
    var items = [];   // { el: w:t node, run: w:r node }
    for (var ri = 0; ri < runs.length; ri++) {
      var ts = runs[ri].getElementsByTagNameNS(W_NS, 't');
      for (var ti = 0; ti < ts.length; ti++) {
        items.push({ el: ts[ti], run: runs[ri] });
      }
    }
    if (items.length === 0) return;

    // Build concatenated text and a character map
    var fullText = '';
    var charMap  = [];   // index → { item index, char offset within that <w:t> }
    for (var ii = 0; ii < items.length; ii++) {
      var txt = items[ii].el.textContent;
      for (var ci = 0; ci < txt.length; ci++) {
        charMap.push({ idx: ii, off: ci });
        fullText += txt[ci];
      }
    }

    // Apply each replacement (may appear more than once)
    var keys = Object.keys(replacements);
    for (var ki = 0; ki < keys.length; ki++) {
      var ph  = keys[ki];
      var val = replacements[ph];
      var pos = fullText.indexOf(ph);

      while (pos !== -1) {
        var endPos   = pos + ph.length;
        var firstIdx = charMap[pos].idx;
        var firstOff = charMap[pos].off;
        var lastIdx  = charMap[endPos - 1].idx;
        var lastOff  = charMap[endPos - 1].off;

        var firstEl = items[firstIdx].el;
        var lastEl  = items[lastIdx].el;

        if (firstIdx === lastIdx) {
          // Placeholder is inside a single <w:t>
          firstEl.textContent =
            firstEl.textContent.substring(0, firstOff) +
            val +
            firstEl.textContent.substring(lastOff + 1);
        } else {
          // Spans multiple <w:t> elements
          var afterLast = lastEl.textContent.substring(lastOff + 1);
          firstEl.textContent = firstEl.textContent.substring(0, firstOff) + val;

          // Clear intermediate <w:t>s
          for (var m = firstIdx + 1; m < lastIdx; m++) {
            items[m].el.textContent = '';
          }
          lastEl.textContent = afterLast;
        }

        // Ensure xml:space="preserve" so leading/trailing spaces survive
        firstEl.setAttribute('xml:space', 'preserve');

        // Rebuild charMap for further replacements
        fullText = '';
        charMap  = [];
        for (var ri2 = 0; ri2 < items.length; ri2++) {
          var t2 = items[ri2].el.textContent;
          for (var ci2 = 0; ci2 < t2.length; ci2++) {
            charMap.push({ idx: ri2, off: ci2 });
            fullText += t2[ci2];
          }
        }

        pos = fullText.indexOf(ph);
      }
    }
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
