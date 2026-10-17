/**
 * 用真实题库文件验证判断题的答案判定口径。
 *
 * 走的是和插件导入完全相同的链路：xlsx → TSV → TableParser.parseToQuestions
 * → normalizeQuestionData，判定也直接调用 main.js 里的真实函数，
 * 不再在这里复制一份逻辑（之前复制过，源码改了测试不会跟着变）。
 *
 * 用法：node tools/check_judge.js [xlsx 文件路径]
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');

const MAIN = path.join(__dirname, '..', 'main.js');
const FILE = process.argv[2] || 'F:/题库100题.xlsx';

/* ---------- 打桩 obsidian 并加载产物 ---------- */
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

const PROBE = `
module.exports.__probe = {
  loadInlineXLSX: typeof loadInlineXLSX === 'function' ? loadInlineXLSX : null,
  TableParser, normalizeQuestionData, getCorrectAnswers, GroupPaperModal,
};
`;
const tmp = path.join(os.tmpdir(), `quiz-bank-judge-${process.pid}.js`);
fs.writeFileSync(tmp, fs.readFileSync(MAIN, 'utf8') + PROBE, 'utf8');
let probe;
try {
  probe = require(tmp).__probe;
} finally {
  try { fs.unlinkSync(tmp); } catch (e) { /* noop */ }
}

const { TableParser, normalizeQuestionData, getCorrectAnswers, GroupPaperModal } = probe;
const XLSX = probe.loadInlineXLSX();
if (!XLSX) { console.error('内联 XLSX 不可用'); process.exit(1); }
if (!fs.existsSync(FILE)) { console.error('找不到题库文件：' + FILE); process.exit(1); }

/* ---------- xlsx → TSV（与 ImportModal 的导入路径一致） ---------- */
const wb = XLSX.read(new Uint8Array(fs.readFileSync(FILE)), { type: 'array', cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false, raw: false });
const tsv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(rows), { FS: '\t', RS: '\n' });

const parsed = TableParser.parseToQuestions(tsv);
const questions = parsed.questions.map(q => normalizeQuestionData(q));

const JUDGE = 'judge';
const judgeQs = questions.filter(q => q.type === JUDGE);
if (judgeQs.length === 0) { console.error('文件里没有判断题'); process.exit(1); }

/* ---------- 判定（直接复用插件里的 judgeAnswerKey） ---------- */
const judgeKey = (q) => GroupPaperModal.prototype.judgeAnswerKey.call({}, q);

let right = 0, wrong = 0, unknown = 0;
const unknowns = [];
for (const q of judgeQs) {
  const k = judgeKey(q);
  if (k === 'wrong') wrong += 1;
  else if (k === 'right') right += 1;
  else { unknown += 1; unknowns.push(q); }
}

console.log(`解析 ${path.basename(FILE)}：${questions.length} 题，其中判断题 ${judgeQs.length} 道`);
console.log(`✓ 正确 ${right} 道，✕ 错误 ${wrong} 道，无法判定 ${unknown} 道`);
if (parsed.errors.length) {
  console.log(`解析告警 ${parsed.errors.length} 条（前 3 条）：`, parsed.errors.slice(0, 3));
}

if (unknown > 0) {
  console.log('无法判定的题（前 5 道）：');
  unknowns.slice(0, 5).forEach(q => {
    console.log('  答案=' + JSON.stringify(q.answer), '|', q.options.map(o => o.key + ':' + o.text).join(' / '),
      '|', String(q.text).slice(0, 30));
  });
}

/* ---------- 一致性校验：答案必须能落到某个选项上 ---------- */
let broken = 0;
for (const q of judgeQs) {
  const answers = getCorrectAnswers(q);
  const keys = q.options.map(o => o.key);
  if (answers.length !== 1 || !keys.includes(answers[0])) broken += 1;
}
if (broken > 0) {
  console.error(`有 ${broken} 道判断题的正确答案落不到任何选项上`);
  process.exit(1);
}
if (right === 0 || wrong === 0) {
  console.error('只识别出一种答案方向，口径可能不匹配');
  process.exit(1);
}
console.log('判定口径正常：两侧都能识别，且答案都落在选项上');
