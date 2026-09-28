/*!
 * build_seed.js — 由 CRM 客户表生成 assets/js/seed-customers.js
 *
 * 刻意复用 store.js 里的 mapCustomerRows()，确保「内置种子数据」与
 * 「页面上导入 CRM 表格」走的是同一套字段映射，不会出现两套口径。
 *
 * 用法：
 *   node tools/build_seed.js "D:/path/to/客户.xlsx"
 */
const fs = require('fs');
const path = require('path');
const { DOMParser } = require('@xmldom/xmldom');

globalThis.DOMParser = DOMParser;
globalThis.window = globalThis;

const ROOT = path.join(__dirname, '..');
require(path.join(ROOT, 'assets/js/store.js'));
const READ = require(path.join(ROOT, 'assets/js/xlsx-read.js'));
const S = globalThis.Store;

async function main() {
  const src = process.argv[2];
  if (!src) {
    console.error('用法: node tools/build_seed.js <客户表.xlsx>');
    process.exit(1);
  }
  if (!fs.existsSync(src)) {
    console.error('文件不存在:', src);
    process.exit(1);
  }

  const buf = fs.readFileSync(src);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const sheet = await READ.read(ab);

  const rows = sheet.rows.filter((r) => r.some((c) => String(c || '').trim()));
  const mapped = S.mapCustomerRows(rows);
  if (mapped.error) {
    console.error('映射失败:', mapped.error);
    process.exit(1);
  }

  // 按名称去重；同时去掉空字段，控制文件体积
  const seen = Object.create(null);
  const out = [];
  let dup = 0;
  for (const rec of mapped.payload) {
    const key = S.customerKey(rec.name);
    if (seen[key]) {
      dup++;
      continue;
    }
    seen[key] = 1;
    const clean = {};
    for (const k of Object.keys(rec)) {
      const v = rec[k];
      if (v !== '' && v != null) clean[k] = v;
    }
    // 统一补一个稳定主键，避免依赖 CRM 的自增 ID
    clean.id = 'c_' + String(out.length + 1).padStart(4, '0');
    out.push(clean);
  }

  const js =
    '/* 由 tools/build_seed.js 自动生成，请勿手工编辑。\n' +
    '   数据来源：' + path.basename(src) + '（' + sheet.sheetName + ' 工作表）\n' +
    '   字段映射与页面上的「导入 CRM 客户表格」完全一致。 */\n' +
    'window.SEED_CUSTOMERS = ' +
    JSON.stringify(out) +
    ';\n';

  const dst = path.join(ROOT, 'assets/js/seed-customers.js');
  fs.writeFileSync(dst, js, 'utf8');

  console.log('工作表      :', sheet.sheetName);
  console.log('读到行数    :', rows.length - 1);
  console.log('去重丢弃    :', dup);
  console.log('写入客户    :', out.length);
  console.log('未识别列    :', (mapped.unmapped || []).length ? mapped.unmapped : '无');
  console.log('输出文件    :', dst, '(' + (fs.statSync(dst).size / 1024).toFixed(1) + ' KB)');
}

main().catch((e) => {
  console.error('生成失败:', e.message);
  process.exit(1);
});
