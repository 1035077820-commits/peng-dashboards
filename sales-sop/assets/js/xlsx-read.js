/*!
 * xlsx-read.js — 极简 .xlsx 读取器（零依赖）
 * 借助浏览器原生 DecompressionStream('deflate-raw') 解压 ZIP 条目，
 * 解析第一个工作表为「行 -> 列」的二维数组，支持 sharedStrings 与内联字符串。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxRead = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SIG_EOCD = 0x06054b50;
  var SIG_CEN = 0x02014b50;

  function findEocd(dv, len) {
    var min = Math.max(0, len - 22 - 65535);
    for (var i = len - 22; i >= min; i--) {
      if (dv.getUint32(i, true) === SIG_EOCD) return i;
    }
    return -1;
  }

  function readCentralDirectory(buf) {
    var dv = new DataView(buf);
    var eocd = findEocd(dv, buf.byteLength);
    if (eocd < 0) throw new Error('不是有效的 xlsx/zip 文件');
    var count = dv.getUint16(eocd + 10, true);
    var offset = dv.getUint32(eocd + 16, true);
    var entries = {};
    var p = offset;
    for (var n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== SIG_CEN) break;
      var entry = {
        method: dv.getUint16(p + 10, true),
        compSize: dv.getUint32(p + 20, true),
        nameLen: dv.getUint16(p + 28, true),
        extraLen: dv.getUint16(p + 30, true),
        commentLen: dv.getUint16(p + 32, true),
        localOffset: dv.getUint32(p + 42, true),
      };
      var name = new TextDecoder().decode(new Uint8Array(buf, p + 46, entry.nameLen));
      entries[name] = entry;
      p += 46 + entry.nameLen + entry.extraLen + entry.commentLen;
    }
    return entries;
  }

  async function extract(buf, entry) {
    var dv = new DataView(buf);
    var nameLen = dv.getUint16(entry.localOffset + 26, true);
    var extraLen = dv.getUint16(entry.localOffset + 28, true);
    var start = entry.localOffset + 30 + nameLen + extraLen;
    var raw = new Uint8Array(buf, start, entry.compSize);
    if (entry.method === 0) return new TextDecoder().decode(raw);
    if (entry.method !== 8) throw new Error('不支持的压缩方式（' + entry.method + '）');
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('当前浏览器不支持解压，请改用 CSV 格式导入');
    }
    var stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return await new Response(stream).text();
  }

  function colToIndex(ref) {
    var m = /^([A-Z]+)/.exec(ref || '');
    if (!m) return -1;
    var s = m[1];
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function textOf(node) {
    var out = '';
    var ts = node.getElementsByTagName('t');
    for (var i = 0; i < ts.length; i++) out += ts[i].textContent;
    return out;
  }

  function parseSharedStrings(xml) {
    if (!xml) return [];
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var items = doc.getElementsByTagName('si');
    var out = [];
    for (var i = 0; i < items.length; i++) out.push(textOf(items[i]));
    return out;
  }

  function parseSheet(xml, shared) {
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var rows = [];
    var rowEls = doc.getElementsByTagName('row');
    for (var i = 0; i < rowEls.length; i++) {
      var rowEl = rowEls[i];
      var r = parseInt(rowEl.getAttribute('r') || String(i + 1), 10);
      var cells = [];
      var cellEls = rowEl.getElementsByTagName('c');
      for (var j = 0; j < cellEls.length; j++) {
        var c = cellEls[j];
        var ci = colToIndex(c.getAttribute('r'));
        if (ci < 0) ci = j;
        var t = c.getAttribute('t');
        var val = '';
        if (t === 's') {
          var vEl = c.getElementsByTagName('v')[0];
          val = vEl ? shared[parseInt(vEl.textContent, 10)] || '' : '';
        } else if (t === 'inlineStr') {
          var isEl = c.getElementsByTagName('is')[0];
          val = isEl ? textOf(isEl) : '';
        } else {
          var v2 = c.getElementsByTagName('v')[0];
          val = v2 ? v2.textContent : '';
        }
        cells[ci] = String(val == null ? '' : val);
      }
      // 补齐空列
      var max = -1;
      for (var k in cells) if (+k > max) max = +k;
      var arr = [];
      for (var m = 0; m <= max; m++) arr.push(cells[m] == null ? '' : cells[m]);
      rows[r - 1] = arr;
    }
    // 补齐空洞
    var out = [];
    for (var x = 0; x < rows.length; x++) out.push(rows[x] || []);
    return out;
  }

  /**
   * 读取 xlsx 的第一个工作表。
   * @param {ArrayBuffer} buf
   * @returns {Promise<{sheetName:string, rows:Array<Array<string>>}>}
   */
  async function read(buf) {
    var entries = readCentralDirectory(buf);

    var workbookXml = entries['xl/workbook.xml'] ? await extract(buf, entries['xl/workbook.xml']) : '';
    var sheetName = 'Sheet1';
    var sheetPath = 'xl/worksheets/sheet1.xml';

    if (workbookXml) {
      var wb = new DOMParser().parseFromString(workbookXml, 'application/xml');
      var sheetEls = wb.getElementsByTagName('sheet');
      if (sheetEls.length) {
        sheetName = sheetEls[0].getAttribute('name') || sheetName;
        var relId =
          sheetEls[0].getAttribute('r:id') ||
          sheetEls[0].getAttributeNS(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'id'
          );
        if (relId && entries['xl/_rels/workbook.xml.rels']) {
          var relXml = await extract(buf, entries['xl/_rels/workbook.xml.rels']);
          var relDoc = new DOMParser().parseFromString(relXml, 'application/xml');
          var rels = relDoc.getElementsByTagName('Relationship');
          for (var i = 0; i < rels.length; i++) {
            if (rels[i].getAttribute('Id') === relId) {
              var target = rels[i].getAttribute('Target') || '';
              target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
              sheetPath = 'xl/' + target;
              break;
            }
          }
        }
      }
    }

    if (!entries[sheetPath]) {
      var found = Object.keys(entries).find(function (k) {
        return /^xl\/worksheets\/sheet\d*\.xml$/.test(k);
      });
      if (!found) throw new Error('文件里没有找到工作表');
      sheetPath = found;
    }

    var shared = entries['xl/sharedStrings.xml']
      ? parseSharedStrings(await extract(buf, entries['xl/sharedStrings.xml']))
      : [];

    var sheetXml = await extract(buf, entries[sheetPath]);
    return { sheetName: sheetName, rows: parseSheet(sheetXml, shared) };
  }

  /** 简单的 CSV 解析（支持引号、逗号、换行） */
  function parseCsv(text) {
    text = text.replace(/^\uFEFF/, '');
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i++;
          } else inQuotes = false;
        } else field += ch;
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        row.push(field);
        field = '';
      } else if (ch === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
      } else if (ch !== '\r') {
        field += ch;
      }
    }
    if (field !== '' || row.length) {
      row.push(field);
      rows.push(row);
    }
    return rows;
  }

  /** 统一入口：按扩展名选择解析方式 */
  async function readFile(file) {
    var name = (file.name || '').toLowerCase();
    if (name.endsWith('.csv') || name.endsWith('.txt')) {
      return { sheetName: 'CSV', rows: parseCsv(await file.text()) };
    }
    return read(await file.arrayBuffer());
  }

  return { read: read, readFile: readFile, parseCsv: parseCsv };
});
