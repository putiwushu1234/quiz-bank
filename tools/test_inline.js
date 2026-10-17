/**
 * 验证内联进 main.js 的 SheetJS 与题库模板真的能用。
 *
 * 做法：给 require('obsidian') 打桩，把 main.js 复制到临时文件并追加一个探针
 * 导出（暴露 loadInlineXLSX 与 TEMPLATE_XLSX_BASE64），然后 require 这份副本，
 * 直接调用产物里真实的加载函数 —— 不再靠解析字符串常量判断，
 * 避免测试脚本和产品代码各写一套逻辑。
 *
 * 用法：node tools/test_inline.js [xlsx 文件路径]
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');

const MAIN = path.join(__dirname, '..', 'main.js');
const XLSX_FILE = process.argv[2] || 'F:/题库100题.xlsx';

/* ---------- 1. 打桩 obsidian ---------- */
const obsidianStub = {
  Plugin: class { constructor() {} async loadData() { return {}; } async saveData() {} },
  PluginSettingTab: class { constructor() {} },
  ItemView: class { constructor() {} },
  Modal: class { constructor() {} },
  Setting: class { constructor() {} },
  Notice: class { constructor() {} },
};
const origLoad = Module._load;
Module._load = function (request, parent, isMain) {
  if (request === 'obsidian') return obsidianStub;
  return origLoad.call(this, request, parent, isMain);
};

/* ---------- 2. 生成带探针的临时副本 ---------- */
const PROBE = `
module.exports.__probe = {
  loadInlineXLSX: typeof loadInlineXLSX === 'function' ? loadInlineXLSX : null,
  template: typeof TEMPLATE_XLSX_BASE64 === 'string' ? TEMPLATE_XLSX_BASE64 : null,
};
`;
const tmp = path.join(os.tmpdir(), `quiz-bank-inline-${process.pid}.js`);
fs.writeFileSync(tmp, fs.readFileSync(MAIN, 'utf8') + PROBE, 'utf8');

let mod, probe;
try {
  mod = require(tmp);
  probe = mod && mod.__probe;
} catch (e) {
  console.error('加载 main.js 失败：', e.message);
  process.exit(1);
} finally {
  try { fs.unlinkSync(tmp); } catch (e) { /* noop */ }
}

/* ---------- 3. 内联的 SheetJS ---------- */
if (!probe || typeof probe.loadInlineXLSX !== 'function') {
  console.error('产物里没有 loadInlineXLSX，内联构建可能未执行');
  process.exit(1);
}

const picked = probe.loadInlineXLSX();
const isUsable = (o) => !!(o && typeof o.read === 'function' && o.utils && typeof o.utils.sheet_to_json === 'function');
if (!isUsable(picked)) {
  console.error('loadInlineXLSX() 没有返回可用的 XLSX 实例：', picked && Object.keys(picked).slice(0, 8));
  process.exit(1);
}
console.log('内联 XLSX 可用：read=' + typeof picked.read + '，版本 ' + picked.version);

// 库不能污染 main.js 自身的导出（导出的应该是插件类，不是 XLSX 实例）
if (typeof mod !== 'function' || mod === picked) {
  console.error('模块导出被内联库污染：exports 应为插件类，实际为 ' + typeof mod);
  process.exit(1);
}

/* ---------- 4. 解析真实题库文件 ---------- */
if (fs.existsSync(XLSX_FILE)) {
  const buf = fs.readFileSync(XLSX_FILE);
  const wb = picked.read(buf, { type: 'buffer' });
  const name = wb.SheetNames[0];
  const rows = picked.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
  console.log(`解析 ${path.basename(XLSX_FILE)}：sheet=${name}，行数=${rows.length}`);
  console.log('首行：', JSON.stringify(rows[0]));
  if (rows.length < 2) {
    console.error('解析结果行数不足，Excel 解析可能异常');
    process.exit(1);
  }
} else {
  console.log('跳过文件解析测试（找不到 ' + XLSX_FILE + '）');
}

/* ---------- 5. 内置模板：完整复现 downloadTemplate() 的 atob -> Uint8Array 路径 ---------- */
if (!probe.template) {
  console.log('跳过模板测试（产物里没有 TEMPLATE_XLSX_BASE64）');
} else {
  const bin = atob(probe.template);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  console.log('内置模板解码：', arr.buffer.byteLength, '字节');

  const wb2 = picked.read(new Uint8Array(arr.buffer), { type: 'array' });
  const rows2 = picked.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]], { header: 1 });
  console.log('内置模板解析：sheet=' + wb2.SheetNames[0] + '，行数=' + rows2.length);
  console.log('首行：', JSON.stringify(rows2[0]));
  if (!rows2.length) {
    console.error('内置模板解析结果为空');
    process.exit(1);
  }
}

console.log('全部通过');
