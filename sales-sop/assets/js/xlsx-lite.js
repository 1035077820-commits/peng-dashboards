/*!
 * xlsx-lite — 极简 .xlsx 生成器（零依赖）
 * 直接手写 OOXML + ZIP(store)，产出真实的 .xlsx 文件，支持中文、换行、列宽、样式。
 * 同时兼容浏览器（全局 XlsxLite）与 Node（module.exports）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxLite = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------------- CRC32 ---------------- */
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xffffffff;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  /* ---------------- ZIP (store, 不压缩) ---------------- */
  var encoder = new TextEncoder();

  function zipStore(files) {
    var parts = [];
    var central = [];
    var offset = 0;

    for (var i = 0; i < files.length; i++) {
      var name = encoder.encode(files[i].name);
      var data = files[i].data;
      var crc = crc32(data);
      var size = data.length;

      var local = new Uint8Array(30 + name.length);
      var lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034b50, true); // local file header
      lv.setUint16(4, 20, true); // version needed
      lv.setUint16(6, 0x0800, true); // flags: UTF-8 文件名
      lv.setUint16(8, 0, true); // method: store
      lv.setUint16(10, 0, true); // mod time
      lv.setUint16(12, 0x2821, true); // mod date (2020-01-01)
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, name.length, true);
      lv.setUint16(28, 0, true);
      local.set(name, 30);

      parts.push(local, data);

      var cen = new Uint8Array(46 + name.length);
      var cv = new DataView(cen.buffer);
      cv.setUint32(0, 0x02014b50, true); // central dir header
      cv.setUint16(4, 20, true); // version made by
      cv.setUint16(6, 20, true); // version needed
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x2821, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, name.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      cen.set(name, 46);
      central.push(cen);

      offset += local.length + size;
    }

    var centralSize = 0;
    for (var j = 0; j < central.length; j++) centralSize += central[j].length;

    var eocd = new Uint8Array(22);
    var ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    var total = offset + centralSize + 22;
    var out = new Uint8Array(total);
    var pos = 0;
    for (var p = 0; p < parts.length; p++) {
      out.set(parts[p], pos);
      pos += parts[p].length;
    }
    for (var q = 0; q < central.length; q++) {
      out.set(central[q], pos);
      pos += central[q].length;
    }
    out.set(eocd, pos);
    return out;
  }

  /* ---------------- XML 工具 ---------------- */
  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
      // 去掉 XML 1.0 非法控制字符
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function colName(n) {
    var s = '';
    n += 1;
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function xmlFile(s) {
    return encoder.encode(s);
  }

  /* ---------------- 样式定义 ---------------- */
  // s 值对应 cellXfs 下标
  // 0 默认 | 1 表头 | 2 自动换行(顶部对齐) | 3 分组标题 | 4 居中
  var STYLES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3">' +
    '<font><sz val="11"/><color theme="1"/><name val="HarmonyOS Sans SC"/><charset val="134"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="HarmonyOS Sans SC"/><charset val="134"/></font>' +
    '<font><b/><sz val="11"/><color theme="1"/><name val="HarmonyOS Sans SC"/><charset val="134"/></font>' +
    '</fonts>' +
    '<fills count="4">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1120BF"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF1F2FB"/><bgColor indexed="64"/></patternFill></fill>' +
    '</fills>' +
    '<borders count="2">' +
    '<border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border>' +
    '<left style="thin"><color rgb="FFD3D6E6"/></left>' +
    '<right style="thin"><color rgb="FFD3D6E6"/></right>' +
    '<top style="thin"><color rgb="FFD3D6E6"/></top>' +
    '<bottom style="thin"><color rgb="FFD3D6E6"/></bottom>' +
    '<diagonal/>' +
    '</border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="5">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
    '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +
    '<alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
    '<alignment horizontal="left" vertical="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +
    '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';

  /**
   * 生成 xlsx。
   * @param {Object} opt
   *   sheetName {string} 工作表名（默认 Sheet1，≤31字符，不含 []:*?/\）
   *   columns   {Array<{header:string,width:number,style?:number}>}
   *   rows      {Array<Array<string|number|{v:*,s:number}>>}
   *   headerStyle {number} 表头样式，默认 1
   *   freeze    {boolean} 是否冻结表头
   * @returns {Uint8Array}
   */
  function build(opt) {
    opt = opt || {};
    var sheetName = String(opt.sheetName || 'Sheet1')
      .replace(/[\[\]\*\?\/\\:]/g, ' ')
      .slice(0, 31) || 'Sheet1';
    var columns = opt.columns || [];
    var rows = opt.rows || [];
    var headerStyle = opt.headerStyle == null ? 1 : opt.headerStyle;
    var colCount = columns.length;

    // 列宽
    var colsXml = '';
    if (columns.length) {
      var c = '';
      for (var i = 0; i < columns.length; i++) {
        var w = columns[i].width || 12;
        c += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
      }
      colsXml = '<cols>' + c + '</cols>';
    }

    // 表头
    var body = '';
    var r = 1;
    if (columns.length && !opt.noHeader) {
      var hc = '';
      for (var h = 0; h < columns.length; h++) {
        hc +=
          '<c r="' + colName(h) + r + '" s="' + headerStyle + '" t="inlineStr"><is><t xml:space="preserve">' +
          esc(columns[h].header == null ? '' : columns[h].header) +
          '</t></is></c>';
      }
      body += '<row r="' + r + '" ht="26" customHeight="1">' + hc + '</row>';
      r++;
    }

    // 数据行
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri] || [];
      var rc = '';
      var maxLines = 1;
      for (var ci = 0; ci < colCount; ci++) {
        var cell = row[ci];
        if (cell == null) continue;
        var val, st;
        if (typeof cell === 'object') {
          val = cell.v;
          st = cell.s == null ? 2 : cell.s;
        } else {
          val = cell;
          st = 2;
        }
        if (val == null || val === '') continue;
        var key = 'r="' + colName(ci) + r + '" s="' + st + '"';
        if (typeof val === 'number' && isFinite(val)) {
          rc += '<c ' + key + '><v>' + val + '</v></c>';
        } else {
          var text = String(val);
          var lines = text.split('\n').length;
          if (lines > maxLines) maxLines = lines;
          rc += '<c ' + key + ' t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
        }
      }
      if (rc) {
        // 依据换行数估算行高，保证多行内容完整显示
        var ht = maxLines > 1 ? maxLines * 15 + 6 : 20;
        body += '<row r="' + r + '" ht="' + ht + '" customHeight="1">' + rc + '</row>';
      }
      r++;
    }

    var lastRow = Math.max(r - 1, 1);
    var dim = colCount ? 'A1:' + colName(colCount - 1) + lastRow : 'A1';
    var pane =
      opt.freeze && colCount
        ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
          '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        : '';

    var sheet =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="' + dim + '"/>' +
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' + pane + '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="16.5"/>' +
      colsXml +
      '<sheetData>' + body + '</sheetData>' +
      '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '</worksheet>';

    var contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';

    var rels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var workbook =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>';

    var workbookRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    return zipStore([
      { name: '[Content_Types].xml', data: xmlFile(contentTypes) },
      { name: '_rels/.rels', data: xmlFile(rels) },
      { name: 'xl/workbook.xml', data: xmlFile(workbook) },
      { name: 'xl/_rels/workbook.xml.rels', data: xmlFile(workbookRels) },
      { name: 'xl/styles.xml', data: xmlFile(STYLES) },
      { name: 'xl/worksheets/sheet1.xml', data: xmlFile(sheet) },
    ]);
  }

  /** 触发浏览器下载 */
  function download(bytes, filename, mime) {
    var blob = new Blob([bytes], {
      type: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 2000);
  }

  return { build: build, download: download, STYLE: { DEFAULT: 0, HEADER: 1, WRAP: 2, GROUP: 3, CENTER: 4 } };
});
