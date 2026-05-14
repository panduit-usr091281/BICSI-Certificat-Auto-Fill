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
  var templateInfo   = null;   // { bgDataUrl, cec, instructor, eventId }

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

      // Extract template metadata (background image, static field values)
      extractTemplateInfo(templateBuffer).then(function (info) {
        templateInfo = info;
      }).catch(function (err) {
        console.warn('Could not extract template info:', err);
        templateInfo = { bgDataUrl: null, cec: '', instructor: '', eventId: '' };
      });

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
     4. TEMPLATE INFO EXTRACTION
     Extracts the background image and static field values (instructor,
     CEC count, event ID) from the uploaded DOCX template so they can
     be used in the HTML-rendered PDF certificates.
     ================================================================== */

  async function extractTemplateInfo(templateBuf) {
    var zip = await JSZip.loadAsync(templateBuf);
    var info = { bgDataUrl: null, cec: '', instructor: '', eventId: '' };

    // ---- background image ----
    var imgFile = zip.file('word/media/image1.png');
    if (imgFile) {
      var imgBlob = await imgFile.async('blob');
      info.bgDataUrl = await compressImageToDataUrl(imgBlob, 1650, 1275, 0.85);
    }

    // ---- SDT tag values + MC block text ----
    var docFile = zip.file('word/document.xml');
    if (!docFile) return info;
    var xml = await docFile.async('string');

    // CEC from its SDT
    var cecMatch = xml.match(/<w:sdt[\s>][\s\S]*?<w:tag\s+w:val="CEC"[\s\S]*?<\/w:sdt>/);
    if (cecMatch) {
      var cecTexts = [];
      cecMatch[0].replace(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g, function (_, t) { cecTexts.push(t); });
      info.cec = cecTexts.join('').trim() || '';
    }

    // Instructor name & event ID from mc:AlternateContent blocks
    var mcBlocks = xml.match(/<mc:AlternateContent[\s>][\s\S]*?<\/mc:AlternateContent>/g) || [];
    for (var i = 0; i < mcBlocks.length; i++) {
      var texts = [];
      mcBlocks[i].replace(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g, function (_, t) { texts.push(t); });
      var joined = texts.join('');

      if (!info.instructor && /Instructor/i.test(joined) && /Signature/i.test(joined)) {
        var instrMatch = joined.match(/Signature:\s*(.+?)(?:\s{2,}|Instructor|$)/i);
        if (instrMatch) info.instructor = instrMatch[1].trim();
      }

      if (!info.eventId) {
        var eidMatch = joined.match(/([A-Z]{2,}-[A-Z]+-[A-Z]+-\d{4}-\d+)/);
        if (eidMatch) info.eventId = eidMatch[1];
      }
    }

    return info;
  }

  /** Compress an image Blob to a JPEG data-URL at the given max size. */
  function compressImageToDataUrl(blob, maxW, maxH, quality) {
    return new Promise(function (resolve) {
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () {
        var scale = Math.min(maxW / img.width, maxH / img.height, 1);
        var canvas = document.createElement('canvas');
        canvas.width  = Math.round(img.width  * scale);
        canvas.height = Math.round(img.height * scale);
        var ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        resolve(null);
      };
      img.src = url;
    });
  }

  /* ==================================================================
     4b. PDF GENERATION (HTML certificate → html2canvas → jsPDF)
     Builds an HTML replica of the certificate using the template's own
     background image and positions the dynamic fields on top.
     ================================================================== */

  /**
   * Render one certificate as a PDF.
   * @param {Object} fields  { name, organization, course, date, instructor, cec, eventId }
   * @returns {Promise<ArrayBuffer>} PDF bytes
   */
  async function renderPdfFromHtml(fields) {
    var container = renderContainer;
    container.innerHTML = '';

    var W = 1100, H = 850;  // landscape letter proportions

    var cert = document.createElement('div');
    cert.style.cssText = 'width:' + W + 'px;height:' + H + 'px;position:relative;overflow:hidden;' +
      'font-family:Calibri,Arial,sans-serif;background:#fff;';

    // Background image (extracted from template DOCX)
    if (templateInfo && templateInfo.bgDataUrl) {
      var bgImg = document.createElement('img');
      bgImg.src = templateInfo.bgDataUrl;
      bgImg.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;';
      cert.appendChild(bgImg);
    }

    // Dynamic text overlays
    addOverlay(cert, fields.name, {
      top: '38%', left: '50%', transform: 'translate(-50%,0)',
      fontSize: '26px', fontWeight: 'bold', color: '#1a1a1a',
      textAlign: 'center', whiteSpace: 'nowrap'
    });

    addOverlay(cert, fields.organization, {
      top: '51%', left: '50%', transform: 'translate(-50%,0)',
      fontSize: '19px', fontStyle: 'italic', color: '#1a1a1a',
      textAlign: 'center', whiteSpace: 'nowrap'
    });

    addOverlay(cert, fields.course, {
      top: '63.5%', left: '50%', transform: 'translate(-50%,0)',
      fontSize: '19px', color: '#1a1a1a', textAlign: 'center',
      maxWidth: '70%', lineHeight: '1.3'
    });

    // Bottom info section
    var bottomData = [
      { label: 'Instructor Signature', value: fields.instructor || '' },
      { label: 'CECs',                 value: fields.cec        || '' },
      { label: 'Event ID',             value: fields.eventId    || '' },
      { label: 'Date of Attendance',   value: fields.date       || '' }
    ];

    var bottom = document.createElement('div');
    bottom.style.cssText = 'position:absolute;bottom:7.5%;left:10%;right:10%;' +
      'display:flex;justify-content:space-between;align-items:flex-end;';

    bottomData.forEach(function (item) {
      var col = document.createElement('div');
      col.style.cssText = 'text-align:center;flex:1;margin:0 6px;';

      var val = document.createElement('div');
      val.style.cssText = 'font-size:12px;font-weight:bold;color:#1a1a1a;' +
        'padding-bottom:3px;border-bottom:1px solid #333;margin-bottom:2px;min-width:90px;';
      val.textContent = item.value;

      var lbl = document.createElement('div');
      lbl.style.cssText = 'font-size:9px;color:#444;letter-spacing:.3px;';
      lbl.textContent = item.label;

      col.appendChild(val);
      col.appendChild(lbl);
      bottom.appendChild(col);
    });

    cert.appendChild(bottom);
    container.appendChild(cert);

    // Allow background image to render
    await delay(400);

    var canvas = await html2canvas(cert, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      logging: false,
      backgroundColor: '#ffffff'
    });

    var jsPDFLib = window.jspdf.jsPDF;
    var pdf = new jsPDFLib({ orientation: 'landscape', unit: 'in', format: 'letter' });
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, 11, 8.5);

    container.innerHTML = '';
    return pdf.output('arraybuffer');
  }

  function addOverlay(parent, text, styles) {
    var el = document.createElement('div');
    el.style.position = 'absolute';
    for (var key in styles) { el.style[key] = styles[key]; }
    el.textContent = text || '';
    parent.appendChild(el);
  }

  /* ==================================================================
     5. MAIN GENERATION FLOW
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

        // ---- DOCX (pixel-perfect from template) ----
        setProgress(i + 1, total, 'Filling template for ' + name + '…');
        var filledDocx = await fillTemplate(templateBuffer, fieldMap);
        outZip.file(baseName + '.docx', filledDocx);

        // ---- PDF (HTML-rendered) ----
        setProgress(i + 1, total, 'Rendering PDF for ' + name + '…');
        var pdfFields = {
          name:         name,
          organization: org,
          course:       course,
          date:         dateStr,
          instructor:   templateInfo ? templateInfo.instructor : '',
          cec:          templateInfo ? templateInfo.cec        : '',
          eventId:      templateInfo ? templateInfo.eventId    : ''
        };
        var pdfBuf = await renderPdfFromHtml(pdfFields);
        outZip.file(baseName + '.pdf', pdfBuf);

        generated++;
      }

      setProgress(total, total, 'Packaging certificates…');
      var blob = await outZip.generateAsync({ type: 'blob' });

      resultSummary.textContent =
        generated + ' certificate(s) generated (' + (generated * 2) + ' files: DOCX + PDF)' +
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
