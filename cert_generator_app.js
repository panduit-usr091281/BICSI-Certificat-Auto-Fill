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
  var courseCol         = document.getElementById('course-col');

  var generateSection  = document.getElementById('step-generate');
  var btnGenerate      = document.getElementById('btn-generate');
  var progressArea     = document.getElementById('progress-area');
  var progressFill     = document.getElementById('progress-fill');
  var statusText       = document.getElementById('status-text');

  var downloadSection  = document.getElementById('step-download');
  var resultSummary    = document.getElementById('result-summary');
  var btnDownload      = document.getElementById('btn-download');

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
    [nameCol, orgCol, dateCol, courseCol].forEach(function (sel) {
      sel.innerHTML = '';
      cols.forEach(function (c) {
        var opt = document.createElement('option');
        opt.value = c;
        opt.textContent = c;
        sel.appendChild(opt);
      });
    });
    autoDetect(nameCol,   cols, ['name', 'full name', 'student', 'participant', 'attendee']);
    autoDetect(orgCol,    cols, ['organization', 'org', 'company', 'employer']);
    autoDetect(dateCol,   cols, ['date', 'attendance', 'date of attendance', 'completion']);
    autoDetect(courseCol,  cols, ['course', 'class', 'training', 'program', 'session']);
  }

  function autoDetect(selectEl, cols, keywords) {
    for (var i = 0; i < cols.length; i++) {
      var low = cols[i].toLowerCase();
      for (var k = 0; k < keywords.length; k++) {
        if (low.indexOf(keywords[k]) !== -1) { selectEl.value = cols[i]; return; }
      }
    }
  }



  /* ==================================================================
     3. DOCX TEMPLATE FILLING
     The BICSI template uses Word Content Controls (Structured Document
     Tags — <w:sdt>) identified by <w:tag w:val="…"/>.  We locate each
     SDT block by its tag name and replace ALL <w:t> text inside the
     content body (<w:sdtContent>) with the new value.
     This approach never parses or re-serialises the XML, so every
     text box, line, shape, image, and drawing is preserved byte-for-byte.
     ================================================================== */

  /**
   * Fill the template with one row of data.
   * `fieldMap` is { tagName: newValue } e.g.
   *   { StudentNameFirstLast: "John Doe", CompanyName: "Panduit", … }
   */
  async function fillTemplate(templateBuf, fieldMap) {
    var zip = await JSZip.loadAsync(templateBuf);

    // Process every XML part that may contain content controls
    var parts = ['word/document.xml'];
    zip.forEach(function (path) {
      if (/^word\/(header|footer)\d*\.xml$/.test(path)) parts.push(path);
    });

    for (var i = 0; i < parts.length; i++) {
      var f = zip.file(parts[i]);
      if (!f) continue;
      var xml = await f.async('string');
      xml = fillContentControls(xml, fieldMap);
      zip.file(parts[i], xml);
    }

    return zip.generateAsync({ type: 'arraybuffer' });
  }

  /** Escape special XML characters in replacement values. */
  function escapeXml(str) {
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');
  }

  /**
   * Find every <w:sdt>…</w:sdt> block whose <w:tag w:val="TAG"/>
   * matches a key in fieldMap, and replace the text inside its
   * <w:sdtContent> with the corresponding value.
   *
   * Strategy for each matched SDT:
   *  - Keep the FIRST <w:t> element and set its text to the new value.
   *  - Empty every OTHER <w:t> so the value isn't duplicated.
   *  - Everything else (run properties, formatting, structure) is untouched.
   */
  function fillContentControls(xmlStr, fieldMap) {
    // Match each <w:sdt>…</w:sdt> (non-greedy, handles nesting via counting)
    var sdtRe = /<w:sdt[\s>][\s\S]*?<\/w:sdt>/g;

    return xmlStr.replace(sdtRe, function (sdtBlock) {
      // Which tag does this SDT have?
      var tagMatch = sdtBlock.match(/<w:tag\s+w:val="([^"]*)"/);
      if (!tagMatch) return sdtBlock;   // no tag — leave alone

      var tagName = tagMatch[1];
      if (!(tagName in fieldMap)) return sdtBlock;  // not a field we care about

      var newValue = escapeXml(fieldMap[tagName]);

      // Find the <w:sdtContent>…</w:sdtContent> portion
      var contentStart = sdtBlock.indexOf('<w:sdtContent');
      if (contentStart === -1) return sdtBlock;

      var beforeContent = sdtBlock.substring(0, contentStart);
      var contentAndAfter = sdtBlock.substring(contentStart);

      // Replace text in <w:t> elements inside the content section
      var isFirst = true;
      var replaced = contentAndAfter.replace(
        /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g,
        function (full, open, _text, close) {
          if (isFirst) {
            isFirst = false;
            // Ensure xml:space="preserve" so spaces aren't trimmed
            if (open.indexOf('xml:space') === -1) {
              open = open.replace('>', ' xml:space="preserve">');
            }
            return open + newValue + close;
          }
          // Empty subsequent <w:t> elements
          return open + close;
        }
      );

      return beforeContent + replaced;
    });
  }

  /* ==================================================================
     4. MAIN GENERATION FLOW
     ================================================================== */

  btnGenerate.addEventListener('click', generateCertificates);

  async function generateCertificates() {
    var nCol     = nameCol.value;
    var oCol     = orgCol.value;
    var dCol     = dateCol.value;
    var cCol     = courseCol.value;

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
        var course = String(row[cCol] != null ? row[cCol] : '').trim();

        if (!name) { skipped++; setProgress(i + 1, total, 'Skipping empty row…'); continue; }

        var dateStr = formatDateValue(dateVal);

        // Map SDT tag names → values from the spreadsheet row
        var fieldMap = {
          StudentNameFirstLast: name,
          CompanyName:          org,
          IssueDate:            dateStr,
          CourseName:           course
        };

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
        var filledDocx = await fillTemplate(templateBuffer, fieldMap);
        outZip.file(baseName + '.docx', filledDocx);

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
     5. UTILITIES
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

})();
