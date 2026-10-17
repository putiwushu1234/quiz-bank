/**
 * 核心逻辑审计脚本（不依赖 Obsidian 运行时）
 *
 * 做法：给 require('obsidian') 打桩，把 main.src.js 复制到临时文件并追加导出块，
 * 然后直接调用源码里的真实函数做边界用例验证 —— 避免测试脚本复制一份逻辑、
 * 结果和产品代码各写一套（check_judge.js 之前就有这个隐患）。
 *
 * 用法：node tools/audit_core.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');

const SRC = path.join(__dirname, 'main.src.js');

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

/* ---------- 2. 生成带导出块的临时副本 ---------- */
const EXPORTS = `
module.exports.__audit = {
  normalizeQuestionData, getCorrectAnswers, normalizeAnswerLetters,
  normalizeQuestionText, sanitizeQuestion, sanitizeBankQuestions, isValidBank,
  TableParser, SimilarityService, generateId, isEmptyValue, progressKeyForLabel,
  extractAnalysisField, formatQuestionAsCallout, ANALYSIS_COLUMN_KEYS,
  QUESTION_TYPES, TYPE_LABELS, QuizView, GroupPaperModal, isQuestionMode, isAndroid
};
`;

const tmp = path.join(os.tmpdir(), `quiz-bank-audit-${process.pid}.js`);
fs.writeFileSync(tmp, fs.readFileSync(SRC, 'utf8') + EXPORTS, 'utf8');

let audit;
try {
  audit = require(tmp).__audit;
} catch (e) {
  console.error('加载源码失败：', e.message);
  process.exit(1);
} finally {
  try { fs.unlinkSync(tmp); } catch (e) { /* noop */ }
}

const {
  normalizeQuestionData, getCorrectAnswers, TableParser,
  sanitizeQuestion, isValidBank, QUESTION_TYPES, isQuestionMode,
} = audit;

/* ---------- 3. 断言框架 ---------- */
let pass = 0;
const fails = [];
function check(name, cond, detail) {
  if (cond) { pass += 1; return; }
  fails.push(detail ? `${name} —— ${detail}` : name);
}

const J = QUESTION_TYPES.JUDGE;
const S = QUESTION_TYPES.SINGLE;
const M = QUESTION_TYPES.MULTIPLE;

/* ---------- 4. 判断题：答案口径 ---------- */
const judgeCases = [
  ['错误', ['B'], '文字「错误」'],
  ['错', ['B'], '单字「错」'],
  ['×', ['B'], '符号 ×'],
  ['F', ['B'], '字母 F'],
  ['false', ['B'], 'false'],
  ['正确', ['A'], '文字「正确」'],
  ['对', ['A'], '单字「对」'],
  ['√', ['A'], '符号 √'],
  ['T', ['A'], '字母 T'],
  ['A', ['A'], '字母答案 A'],
  ['B', ['B'], '字母答案 B'],
];
for (const [raw, want, label] of judgeCases) {
  const q = normalizeQuestionData({
    type: J, text: '判断题', answer: raw, options: [], extraFields: {},
  });
  const got = getCorrectAnswers(q);
  check(`判断题答案 ${label}`, JSON.stringify(got) === JSON.stringify(want),
    `期望 ${JSON.stringify(want)}，实际 ${JSON.stringify(got)}`);
}

/* 空白 / 无法识别的答案不能变成「永远判对」 */
for (const raw of ['', '   ', '不知道']) {
  const q = normalizeQuestionData({ type: J, text: 'x', answer: raw, options: [] });
  const got = getCorrectAnswers(q);
  check(`判断题空白答案「${raw}」不应被判对`, got.length === 0, `实际 ${JSON.stringify(got)}`);
}

/* ---------- 5. 选择题：答案与选项 key ---------- */
const single = normalizeQuestionData({
  type: S, text: '单选', answer: 'ａ',
  options: [{ key: 'a', text: '选项一' }, { key: 'b', text: '选项二' }],
});
check('单选：全角小写答案规范化', single.answer === 'A', `答案=${single.answer}`);
check('单选：选项 key 规范化为大写',
  single.options.map(o => o.key).join('') === 'AB',
  `keys=${single.options.map(o => o.key).join('')}`);
check('单选：正确答案', getCorrectAnswers(single).join('') === 'A');

const multi = normalizeQuestionData({
  type: M, text: '多选', answer: 'A、B',
  options: [{ key: 'A', text: '一' }, { key: 'B', text: '二' }, { key: 'C', text: '三' }],
});
check('多选：顿号分隔答案', getCorrectAnswers(multi).sort().join('') === 'AB',
  `实际 ${JSON.stringify(getCorrectAnswers(multi))}`);

const dupKey = normalizeQuestionData({
  type: S, text: '重复key', answer: 'A',
  options: [{ key: 'A', text: '一' }, { key: 'A', text: '二' }, { key: 'C', text: '三' }],
});
check('重复选项 key 被重编',
  new Set(dupKey.options.map(o => o.key)).size === 3,
  `keys=${dupKey.options.map(o => o.key).join('')}`);

/* ---------- 6. CSV / TSV 解析 ---------- */
const csv = [
  '题型,题目,A,B,C,D,答案,解析',
  '单选,"含逗号,的题干",甲,乙,丙,丁,B,因为',
  '多选,多选题干,甲,乙,丙,丁,AC,略',
  '判断,判断题干,正确,错误,,,错误,略',
].join('\n');

const qs = TableParser.parseToQuestions(csv).questions;
check('CSV 解析出 3 道题', qs.length === 3, `实际 ${qs.length}`);
check('CSV 引号内的逗号未被拆列', qs[0] && qs[0].text === '含逗号,的题干',
  `题干=${JSON.stringify(qs[0] && qs[0].text)}`);
check('CSV 单选答案=B', qs[0] && qs[0].answer === 'B', `答案=${qs[0] && qs[0].answer}`);
check('CSV 多选答案=AC', qs[1] && qs[1].answer === 'AC', `答案=${qs[1] && qs[1].answer}`);
check('CSV 判断题答案归一为 B', qs[2] && qs[2].answer === 'B',
  `答案=${qs[2] && qs[2].answer}`);
check('CSV 解析列', qs[2] && qs[2].analysis === '略',
  `解析=${JSON.stringify(qs[2] && qs[2].analysis)}`);

/* 制表符分隔 + 跨行字段 */
const tsv = '题目\t答案\n"多行\n题干"\tA';
const tq = TableParser.parseToQuestions(tsv).questions;
check('TSV 自动识别分隔符', tq.length === 1, `实际 ${tq.length}`);
check('引号内换行不被切断', tq[0] && tq[0].text.includes('\n'),
  `题干=${JSON.stringify(tq[0] && tq[0].text)}`);

/* 空内容 / 纯表头：应当给出可读的报错，而不是空对象或崩溃 */
for (const [name, input] of [['空字符串', ''], ['只有表头', '题型,题目,答案'], ['空行', '\n\n\n']]) {
  let msg = null;
  try { TableParser.parseToQuestions(input); } catch (e) { msg = e.message; }
  check(`解析「${name}」给出可读报错`, !!msg, msg ? '未抛错' : '抛错信息为空');
}

/* 缺题干列 / 缺答案列：报错要指明缺的是哪一列 */
for (const [name, input, kw] of [
  ['缺题干列', '答案,选项A\nA,x', '题目'],
  ['缺答案列', '题目,A\n题干,x', '答案'],
]) {
  let msg = null;
  try { TableParser.parseToQuestions(input); } catch (e) { msg = e.message; }
  check(`${name} 的报错包含「${kw}」`, !!msg && msg.includes(kw), msg || '未抛错');
}

/* ---------- 7. 脏数据清洗 ---------- */
check('sanitizeQuestion 丢弃非对象', sanitizeQuestion(null) === null);
check('isValidBank 拒绝缺 questions', isValidBank({ id: 'a', name: 'b' }) === false);
check('isValidBank 接受合法题库',
  isValidBank({ id: 'a', name: 'b', questions: [] }) === true);

const dirtyBank = {
  id: 'x', name: '脏数据', questions: [
    { type: S, text: '正常题', answer: 'A', options: [{ key: 'A', text: '一' }] },
    null,
    { text: '缺 type' },
    { type: S, text: '缺答案', options: [{ key: 'A', text: '一' }] },
  ],
};
let dirtyOk = true;
try { audit.sanitizeBankQuestions(dirtyBank); } catch (e) { dirtyOk = false; }
check('清洗脏题库不抛错', dirtyOk);
check('清洗后题目数合法', Array.isArray(dirtyBank.questions) && dirtyBank.questions.length >= 1,
  `实际 ${dirtyBank.questions && dirtyBank.questions.length}`);

/* ---------- 8. 相似度（错题去重 / 相似题） ---------- */
const sim = audit.SimilarityService.diceCoefficient('单选题干内容', '单选题干内容');
check('相同文本相似度为 1', Math.abs(sim - 1) < 1e-9, `实际 ${sim}`);
const sim2 = audit.SimilarityService.diceCoefficient('完全不同的内容', '');
check('空串相似度为 0', sim2 === 0, `实际 ${sim2}`);

/* ---------- 9. 选项乱序后答案必须跟着走 ---------- */
const { QuizView } = audit;

/** 用固定随机序列洗牌，保证测试可复现 */
function shuffleWith(fn) {
  const orig = Math.random;
  let i = 0;
  const seq = [0, 0, 0, 0, 0];   // 恒取 0 => 每次都和首位交换，等于逆序，必定发生真实换位
  Math.random = () => seq[i++ % seq.length];
  try { return fn(); } finally { Math.random = orig; }
}

const viewStub = { plugin: { settings: { shuffleOptions: true } } };

// 选择题：答案字母必须重映射
const shuffleQ = {
  id: 's1', type: S, text: '单选乱序', answer: 'A',
  options: [{ key: 'A', text: '甲' }, { key: 'B', text: '乙' }, { key: 'C', text: '丙' }],
};
const shuffled = shuffleWith(() => QuizView.prototype.applyOptionShuffle.call(viewStub, [shuffleQ]))[0];
const shuffledCorrect = getCorrectAnswers(shuffled);
const textOfCorrect = shuffled.options
  .filter(o => shuffledCorrect.includes(o.key)).map(o => o.text).join(',');
check('乱序后答案仍指向原正确选项（字母答案）', textOfCorrect === '甲',
  `正确答案指向了「${textOfCorrect}」，应为「甲」；answer=${shuffled.answer}`);

// 判断题不参与乱序：选项固定是「正确 / 错误」两个语义标签，打乱后界面会变成「错误 / 正确」
const shuffleJ = {
  id: 'j1', type: J, text: '判断乱序', answer: 'B',
  options: [{ key: 'A', text: '正确' }, { key: 'B', text: '错误' }],
};
const shuffledJ = shuffleWith(() => QuizView.prototype.applyOptionShuffle.call(viewStub, [shuffleJ]))[0];
check('判断题不参与选项乱序', shuffledJ === shuffleJ
  && shuffledJ.options.map(o => o.text).join(',') === '正确,错误',
  `选项变成了「${shuffledJ.options.map(o => o.text).join(',')}」`);
check('判断题答案不受乱序影响', shuffledJ.answer === 'B', `answer=${shuffledJ.answer}`);

// 未开启乱序时不应产生任何副本或改写
const kept = QuizView.prototype.applyOptionShuffle.call(
  { plugin: { settings: { shuffleOptions: false } } }, [shuffleQ]);
check('未开启乱序时原题对象不被改写', kept[0] === shuffleQ && shuffleQ.options[0].key === 'A');

/* 选项顺序被调换成「错误 / 正确」时，判定不能失效 */
const { GroupPaperModal } = audit;
const judgeKeyOf = (q) => GroupPaperModal.prototype.judgeAnswerKey.call({}, q);
const swapped = normalizeQuestionData({
  type: J, text: '顺序反了的判断题', answer: '错误',
  options: [{ key: 'A', text: '错误' }, { key: 'B', text: '正确' }],
});
check('选项顺序调换后仍能识别「错误」', judgeKeyOf(swapped) === 'wrong',
  `实际 ${judgeKeyOf(swapped)}`);
const swappedRight = normalizeQuestionData({
  type: J, text: '顺序反了的判断题2', answer: '正确',
  options: [{ key: 'A', text: '错误' }, { key: 'B', text: '正确' }],
});
check('选项顺序调换后仍能识别「正确」', judgeKeyOf(swappedRight) === 'right',
  `实际 ${judgeKeyOf(swappedRight)}`);

/* ---------- 9.5 答题页类名契约（CSS 第 33 节依赖 .quiz-question-view） ---------- */
const QUESTION_MODES = ['practice', 'review', 'wrong', 'favorites', 'group'];
const OTHER_MODES = ['list', 'chapter-select', 'search'];
for (const m of QUESTION_MODES) {
  check(`模式 ${m} 标记为答题页`, isQuestionMode(m) === true, `返回 ${isQuestionMode(m)}`);
}
for (const m of OTHER_MODES) {
  check(`模式 ${m} 不标记为答题页`, isQuestionMode(m) === false, `返回 ${isQuestionMode(m)}`);
}
check('未知模式不标记为答题页', isQuestionMode('unknown') === false, '返回 true');

{
  const raw = fs.readFileSync(path.join(__dirname, '..', 'styles.css'), 'utf8');
  const css = raw.replace(/\/\*[\s\S]*?\*\//g, '');   // 去掉注释，只看真实规则
  check('styles.css 含 .quiz-question-view 规则',
    css.includes('.quiz-container.quiz-question-view > .quiz-question-card'),
    '缺少选择器，答题页会铺满整宽');
  check('styles.css 已不再依赖 :has()',
    !css.includes(':has('),
    '仍在使用 :has()，旧版 Android WebView 下会失效');
}

/* ---------- 10. 输出 ---------- */
console.log(`\n通过 ${pass} 项`);
if (fails.length) {
  console.error(`失败 ${fails.length} 项：`);
  for (const f of fails) console.error('  ✗ ' + f);
  process.exit(1);
}
console.log('核心逻辑审计全部通过');
