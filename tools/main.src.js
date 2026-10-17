'use strict';

const { Plugin, PluginSettingTab, ItemView, Modal, Setting, Notice } = require('obsidian');

// v2.7.2: XLSX 库用于直接导入 Excel 文件（.xls / .xlsx）
//
// 加载策略（多来源探测，桌面端与移动端通用）：
//   1) 插件同目录的 xlsx.full.min.js —— 本地开发时可用
//   2) node_modules/xlsx —— 本地 npm 安装时可用
//   3) 裸模块 'xlsx'
//   4) main.js 里内联的副本（见文件末尾的 loadInlineXLSX）—— 官方分发的兜底，
//      因为 Obsidian 只分发 main.js / manifest.json / styles.css
//
// v2.7.4: 改为「首次真正用到 Excel 时才加载」—— 约 0.9MB 的解析库不再拖累插件启动，
// 启动阶段也不会再产生与之相关的告警。
//
// ⚠ 重要：不能只信任 require() 的返回值。
// SheetJS 构建末尾的分支是：
//     if (typeof exports !== "undefined") make_xlsx_lib(exports);
//     else if (module && module.exports) make_xlsx_lib(module.exports);
// Obsidian 的插件加载器里 `exports` 存在，但与 `module.exports` 不是同一对象，
// 于是库把自己挂到了 `exports` 上，而 require() 返回的 `module.exports` 上
// **没有** read/write —— 直接判断 `mod.read` 会误判为"加载失败"。
// 因此每个来源都按 require 返回值 → globalThis → window 的顺序探测。
//
// 全部失败时 XLSX 保持 null，界面会提示改用 CSV，不会抛错崩溃。
let XLSX = null;
let _xlsxRequireTried = false;

/** 判断一个对象是否是可用的 XLSX 实例 */
function isUsableXLSX(o) {
  return !!(o && typeof o.read === 'function' && o.utils && typeof o.utils.sheet_to_json === 'function');
}

/**
 * v2.7.4: 通过 require 探测 XLSX（懒加载：只尝试一次，结果缓存）。
 * 不在插件启动阶段调用 —— 从不导入 Excel 的用户不必为它付出启动开销。
 * @returns {object|null}
 */
function tryRequireXLSX() {
  if (XLSX) return XLSX;
  if (_xlsxRequireTried) return null;
  _xlsxRequireTried = true;

  const candidates = [
    './xlsx.full.min.js',
    './node_modules/xlsx/xlsx.js',
    'xlsx'
  ];

  /** 从一个候选来源里找出可用的 XLSX（模块返回值 / 全局挂载都可能） */
  const probe = (name) => {
    let mod = null;
    try { mod = require(name); } catch (e) { /* 该路径不可用，继续探测 */ }

    // ① require() 返回值本身
    if (isUsableXLSX(mod)) return mod;
    // ② 某些打包器把命名空间挂在 .default 上
    if (mod && isUsableXLSX(mod.default)) return mod.default;
    // ③ 库把自己挂到了全局（不同构建的兜底行为）
    try { if (isUsableXLSX(globalThis.XLSX)) return globalThis.XLSX; } catch (e) { /* noop */ }
    try { if (typeof window !== 'undefined' && isUsableXLSX(window.XLSX)) return window.XLSX; } catch (e) { /* noop */ }
    return null;
  };

  for (const name of candidates) {
    const found = probe(name);
    if (found) { XLSX = found; return XLSX; }
  }
  return null;
}

/**
 * v2.7.6: 确保 Excel 解析能力可用。
 *
 * 原来这里有一段「从插件目录读 xlsx.full.min.js 再动态执行」的兜底，
 * 但 Obsidian 只分发 main.js / manifest.json / styles.css，那个文件在用户端本来就不存在；
 * 而且动态代码求值（eval / Function 构造器）是社区审核清单明确禁止的。
 * 现在整个库在构建时就被编译进 main.js（loadInlineXLSX），
 * 所以这里只需要走一次懒加载探测。
 *
 * @param {object} plugin 插件实例（保留参数以兼容既有调用点）
 * @param {string} pluginDir 插件目录（保留参数以兼容既有调用点）
 * @returns {Promise<object|null>} 可用的 XLSX 实例或 null
 */
async function ensureXLSXLoaded(plugin, pluginDir) {
  if (tryRequireXLSX()) return XLSX;
  console.warn('[QuizBank] Excel 解析库不可用，请改用 CSV（UTF-8 逗号分隔）导入');
  return null;
}

/* ============================================================
   Section 1: 常量与工具函数
   ============================================================ */

const VIEW_TYPE_QUIZ = 'quiz-bank-view';

// 数据文件版本号（导出/导入备份时写入）—— 只在数据格式变化时递增，与 manifest 的插件版本号相互独立
const DATA_VERSION = '2.7.4';

const QUESTION_TYPES = {
  SINGLE: 'single',
  MULTIPLE: 'multiple',
  JUDGE: 'judge'
};

const TYPE_LABELS = {
  [QUESTION_TYPES.SINGLE]: '单选题',
  [QUESTION_TYPES.MULTIPLE]: '多选题',
  [QUESTION_TYPES.JUDGE]: '判断题'
};

/** 安全设置文本内容（防 XSS） */
function safeSetText(el, text) {
  if (el) el.textContent = text || '';
}

/** 防抖 */
function debounce(fn, delay) {
  let timer = null;
  return function (...args) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      fn.apply(this, args);
      timer = null;
    }, delay);
  };
}

/** 生成唯一 ID */
function generateId() {
  // 时间戳 + 两段随机数（约 36^12 组合），避免大规模导入时 ID 碰撞
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 10) + Math.random().toString(36).slice(2, 10);
}

/** 触觉反馈（移动端） */
function hapticFeedback() {
  if (navigator.vibrate) {
    navigator.vibrate(10);
  }
}

/** 判断值是否为"空" */
function isEmptyValue(val) {
  if (!val) return true;
  const trimmed = String(val).trim();
  return trimmed === '' || trimmed === '\\' || trimmed === '-' || trimmed === '无' || trimmed === 'N/A' || trimmed === 'n/a';
}


/**
 * v2.7.4: 进度记忆用的键。
 * 与答题记录键（getPracticeKey）刻意分开：章节练习以前和"刷题模式"共用 practice 槽位，
 * 续做时会跳到完全不同题集的位置上。
 */
function progressKeyForLabel(label) {
  switch (label) {
    case '刷题模式': return 'practice';
    case '章节练习': return 'chapter';
    case '错题重练': return 'wrong';
    case '收藏练习': return 'favorites';
    case '组卷练习': return 'group';
    default: return null;   // '搜索结果' 等一次性题集不记进度
  }
}

/** 只有"题集相对固定"的入口才自动续做，避免错题/收藏/组卷这类动态题集跳错位置 */
const RESUMABLE_PROGRESS_KEYS = new Set(['practice', 'chapter']);

/** v2.5: 题干规范化（用于导入查重：忽略空白与大小写差异） */
function normalizeQuestionText(text) {
  return String(text || '').replace(/\s+/g, '').toLowerCase();
}

/**
 * v2.6.1: 答案字母规范化（修复多选题"选对却判错"）
 * 全角字母→半角、统一大写、剔除「、」/空格/标点/数字等杂字符，只保留 A-Z
 */
function normalizeAnswerLetters(answer) {
  return String(answer || '')
    .replace(/[\uFF21-\uFF3A\uFF41-\uFF5A]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .toUpperCase()
    .replace(/[^A-Z]/g, '');
}

/** v2.7.4: 「解析」类列名（导入时归位到 analysis，而不是当成附加信息） */
const ANALYSIS_COLUMN_KEYS = ['解析', '答案解析', '题目解析', '试题解析', '分析', '详解'];

/**
 * v2.7.4: 从附加字段里取出"解析"，并从 extraFields 中移除。
 * 导入的解析列与编辑弹窗写的 analysis 从此是同一个字段，
 * 不会再出现答题页/复制出的 Callout 里两个"解析"块。
 * @returns {string} 解析内容，没有则返回 ''
 */
function extractAnalysisField(extraFields) {
  if (!extraFields || typeof extraFields !== 'object') return '';
  for (const key of ANALYSIS_COLUMN_KEYS) {
    if (!Object.prototype.hasOwnProperty.call(extraFields, key)) continue;
    const val = extraFields[key];
    delete extraFields[key];
    return isEmptyValue(val) ? '' : String(val);
  }
  return '';
}

/** 判断题：表示「错误」的选项文本（含常见符号与英文写法） */
const JUDGE_WRONG_TEXTS = ['错误', '错', '×', '✕', 'x', 'f', 'false', '否', '不正确'];
/** 判断题：表示「正确」的选项文本 */
const JUDGE_RIGHT_TEXTS = ['正确', '对', '√', '✓', 't', 'true', 'y', 'yes', '是'];

/**
 * v2.7.6: 找出判断题的「错误项 / 正确项」key。
 * 必须按选项文本找，不能默认「第一个是正确、第二个是错误」——
 * 选项顺序被调换过时，那样会把两个方向指向同一个 key。
 */
function judgeOptionKeys(q) {
  const opts = (q && Array.isArray(q.options)) ? q.options : [];
  const text = (o) => String((o && o.text) || '').trim().toLowerCase();
  const wrongOpt = opts.find(o => JUDGE_WRONG_TEXTS.includes(text(o)));
  const rightOpt = opts.find(o => JUDGE_RIGHT_TEXTS.includes(text(o)));
  return {
    wrongKey: wrongOpt ? wrongOpt.key : (opts[1] ? opts[1].key : 'B'),
    rightKey: rightOpt ? rightOpt.key : (opts[0] ? opts[0].key : 'A')
  };
}

/** v2.6.1: 获取规范化后的正确答案字母数组（判定统一入口，兼容脏数据） */
function getCorrectAnswers(q) {
  if (!q) return [];
  if (q.type === QUESTION_TYPES.JUDGE) {
    const half = normalizeAnswerLetters(q.answer);
    if (half === 'A' || half === 'B') return [half];
    const raw = String(q.answer || '').trim();
    const keys = (q.options || []).map(o => o.key);
    // v2.7.4: 补充 Y/Yes/1 与 N/No/0 等常见判断题写法
    if (/^(√|对|是|正确|t|true|y|yes|1)/i.test(raw) && keys[0]) return [keys[0]];
    if (/^(×|✕|x|错|否|不正确|f|false|n|no|0)/i.test(raw) && keys[1]) return [keys[1]];
    return half ? [half[0]] : [];
  }
  const norm = normalizeAnswerLetters(q.answer);
  return norm ? [...new Set(norm.split(''))] : [];
}

/**
 * v2.6.1: 载入时规范化题目数据
 * - 选择题答案：全角→半角、大写、去杂字符、去重
 * - 选项 key：规范化为半角大写字母；重复/非法 key 按位置顺延重编
 * - 判断题答案：兼容 √/×/T/F/对/错 等，统一为 A（正确）/ B（错误）
 */
function normalizeQuestionData(q) {
  if (!q || typeof q !== 'object') return q;

  // v2.7.4: 旧数据的「解析」存放在 extraFields 里，统一迁移到 analysis
  if ((!q.analysis || String(q.analysis).trim() === '') && q.extraFields && typeof q.extraFields === 'object') {
    const moved = extractAnalysisField(q.extraFields);
    if (moved) q.analysis = moved;
  }

  if (q.type === QUESTION_TYPES.JUDGE) {
    // v2.7.4: 判断题没有选项时补齐 A.正确 / B.错误。
    // 否则 getCorrectAnswers 返回空数组，这道题永远判错（编辑题目时清空选项也会触发）。
    if (!Array.isArray(q.options) || q.options.length === 0) {
      q.options = [
        { key: 'A', text: '正确' },
        { key: 'B', text: '错误' }
      ];
    }
    // v2.7.6: 原正则带 $ 锚点，只认单字「错 / 对」，
    // 「错误 / 正确」这类两字写法永远匹配不上，答案会一直停在中文上，
    // 后续所有按选项 key 比对的地方（选项乱序、组卷筛选）都会错位。
    // 这里改为「去空白 + 前缀匹配」，并把否定形态放在最前面判断，
    // 否则「不正确 / 不对 / 不是」会被下面的「对 / 是」前缀误判成正确；
    // 文字答案还要按选项的实际 key 落位，不能想当然地写死成 A / B。
    const raw = String(q.answer || '').trim().replace(/\s+/g, '');
    const half = normalizeAnswerLetters(raw);
    const { wrongKey, rightKey } = judgeOptionKeys(q);
    if (half === 'A' || half === 'B') {
      if (q.answer !== half) q.answer = half;
    } else if (/^(不正确|不对|不是|×|✕|x|错|否|f|false|n|no|0)/i.test(raw)) {
      q.answer = wrongKey;
    } else if (/^(√|对|是|正确|t|true|y|yes|1)/i.test(raw)) {
      q.answer = rightKey;
    }
    return q;
  }

  const ans = normalizeAnswerLetters(q.answer);
  if (ans) q.answer = [...new Set(ans.split(''))].join('');

  if (Array.isArray(q.options) && q.options.length > 0) {
    const used = new Set();
    q.options.forEach(o => {
      let k = typeof o.key === 'string' ? normalizeAnswerLetters(o.key) : '';
      if (!k || used.has(k)) {
        let code = 65;
        while (used.has(String.fromCharCode(code))) code++;
        k = String.fromCharCode(code);
      }
      used.add(k);
      if (o.key !== k) o.key = k;
    });
    // v2.6.1: 若 key 不是 A、B、C… 连续前缀（脏数据，如按选项首字母误解析出的 E/L/C），
    // 按位置重编为标准字母，并按旧字母的位置重映射答案
    const seq = q.options.map((o, i) => String.fromCharCode(65 + i));
    const oldKeys = q.options.map(o => o.key);
    if (oldKeys.some((k, i) => k !== seq[i])) {
      const posMap = {};
      oldKeys.forEach((k, i) => { if (!(k in posMap)) posMap[k] = seq[i]; });
      q.options.forEach((o, i) => { o.key = seq[i]; });
      if (q.answer) {
        q.answer = String(q.answer).split('').map(ch => posMap[ch] || ch).join('');
      }
    }
  }
  return q;
}

/* ============================================================
   复制题目功能：将题目格式化为 Obsidian 可折叠 Callout Markdown
   ============================================================ */

/** 行首 [! 可能被 Obsidian 误识别为 Callout 标记，插入零宽空格转义 */
function escapeCalloutMarker(line) {
  return String(line).replace(/^(\s*)\[!/, '$1\u200B[!');
}

/**
 * 将题目格式化为可折叠 Callout 的 Markdown 文本
 * - 外层 [!quote]+（默认展开）：题干 + 选项
 * - 内层 [!question]-（默认折叠）：答案（统一走 getCorrectAnswers 兼容脏数据）
 * - 内层 [!info]-（默认折叠）：解析 + 额外字段（两者均无内容时整个省略）
 * 只读 question / bankName，不改动任何数据结构
 */
function formatQuestionAsCallout(question, bankName) {
  if (!question) return '';
  const bank = (bankName && String(bankName).trim() !== '') ? String(bankName).trim() : '未知题库';
  const lines = [];

  // 外层 Callout 头（默认展开 +）
  lines.push(`> [!quote]+ 插入题目｜来源：${bank}`);

  // 题干：保留换行，首行加 **题干**： 前缀，每行前加 "> "
  const qText = question.text != null ? String(question.text) : '';
  if (qText.trim() !== '') {
    qText.split(/\r?\n/).forEach((l, i) => {
      lines.push(`> ${i === 0 ? '**题干**：' : ''}${escapeCalloutMarker(l)}`);
    });
  }

  // 选项：按 question.options 顺序输出；判断题无选项时补 A. 正确 / B. 错误
  let options = Array.isArray(question.options) ? question.options : [];
  if (options.length === 0 && question.type === QUESTION_TYPES.JUDGE) {
    options = [{ key: 'A', text: '正确' }, { key: 'B', text: '错误' }];
  }
  options.forEach(opt => {
    const key = (opt && opt.key) ? String(opt.key) : '';
    const text = (opt && opt.text != null) ? String(opt.text) : '';
    text.split(/\r?\n/).forEach((l, i) => {
      lines.push(`> ${i === 0 ? `${key}. ` : ''}${escapeCalloutMarker(l)}`);
    });
  });

  // 答案 Callout（默认折叠 -）
  const answers = getCorrectAnswers(question);
  const answerText = answers.length > 0 ? answers.join('') : '（无答案）';
  lines.push('>');
  lines.push('> > [!question]- 答案');
  answerText.split(/\r?\n/).forEach(l => lines.push(`> > ${escapeCalloutMarker(l)}`));

  // 解析 Callout（默认折叠 -）：解析与额外字段均无内容时整个省略
  const analysis = question.analysis != null ? String(question.analysis).trim() : '';
  const extraFields = (question.extraFields && typeof question.extraFields === 'object')
    ? Object.entries(question.extraFields).filter(([k, v]) => k && !isEmptyValue(v))
    : [];
  if (analysis !== '' || extraFields.length > 0) {
    lines.push('>');
    lines.push('> > [!info]- 解析');
    if (analysis !== '') {
      analysis.split(/\r?\n/).forEach(l => lines.push(`> > ${escapeCalloutMarker(l)}`));
    }
    extraFields.forEach(([k, v]) => {
      String(v).split(/\r?\n/).forEach((l, i) => {
        lines.push(`> > ${i === 0 ? `**${k}**：` : ''}${escapeCalloutMarker(l)}`);
      });
    });
  }

  return lines.join('\n');
}

/**
 * 复制题目为 Callout Markdown 到剪贴板（桌面端/移动端通用）
 * 优先 navigator.clipboard.writeText；失败回退 textarea + execCommand('copy')
 */
async function copyQuestionToClipboard(question, bankName) {
  const text = formatQuestionAsCallout(question, bankName);
  if (!text) {
    new Notice('复制失败：题目为空');
    return false;
  }
  let ok = false;
  // 优先 Clipboard API（均在点击回调内调用，满足用户手势要求）
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      ok = true;
    }
  } catch (e) {
    ok = false;
  }
  // 回退：临时 textarea + execCommand('copy')
  if (!ok) {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      ok = document.execCommand('copy');
      ta.remove();
    } catch (e) {
      ok = false;
    }
  }
  if (ok) new Notice('题目已复制为 Callout 格式');
  else new Notice('复制失败，请检查剪贴板权限');
  return ok;
}

/**
 * v2.7.4: 题库结构校验。载入、迁移、导入备份共用同一套判断，
 * 避免损坏或第三方来源的数据进内存后在 getSearchIndex / renderOptions 里以 TypeError 炸开。
 */
function isValidBank(b) {
  return !!(b && typeof b === 'object'
    && typeof b.id === 'string' && b.id
    && typeof b.name === 'string' && b.name
    && Array.isArray(b.questions));
}

/** 把单题数据补齐成插件内部约定的结构（返回 null 表示该条不可用） */
function sanitizeQuestion(q) {
  if (!q || typeof q !== 'object') return null;
  if (typeof q.id !== 'string' || !q.id) q.id = generateId();
  q.text = q.text == null ? '' : String(q.text);
  q.answer = q.answer == null ? '' : String(q.answer);
  q.analysis = q.analysis == null ? '' : String(q.analysis);
  if (!Array.isArray(q.tags)) q.tags = [];
  if (!q.extraFields || typeof q.extraFields !== 'object') q.extraFields = {};
  if (!Array.isArray(q.chapterPath)) q.chapterPath = null;
  if (!q.type) q.type = QUESTION_TYPES.SINGLE;
  const rawOptions = Array.isArray(q.options) ? q.options : [];
  q.options = rawOptions.map((o, i) => {
    const fallbackKey = String.fromCharCode(65 + i);
    if (o && typeof o === 'object') {
      return { key: o.key == null ? fallbackKey : String(o.key), text: o.text == null ? '' : String(o.text) };
    }
    return { key: fallbackKey, text: o == null ? '' : String(o) };
  });
  return q;
}

/** 清洗整个题库的题目数组（含答案/选项 key 规范化） */
function sanitizeBankQuestions(bank) {
  bank.questions = bank.questions
    .map(q => normalizeQuestionData(sanitizeQuestion(q)))
    .filter(Boolean);
  return bank;
}

/** 章节树构建器：从题目的 chapterPath 数组构建层级树 */
class ChapterTreeBuilder {
  /**
   * @param {Array} questions - 包含 chapterPath 字段的题目数组
   * @returns {Object|null} 章节树根节点，或 null
   */
  static build(questions) {
    const root = { name: '__root__', path: '', children: {}, questionIds: [], level: 0 };
    let hasChapters = false;

    for (const q of questions) {
      if (!q.chapterPath || q.chapterPath.length === 0) continue;
      hasChapters = true;
      let node = root;
      for (let i = 0; i < q.chapterPath.length; i++) {
        const seg = q.chapterPath[i];
        if (!node.children[seg]) {
          node.children[seg] = {
            name: seg,
            path: q.chapterPath.slice(0, i + 1).join(' > '),
            children: {},
            questionIds: [],
            level: i + 1
          };
        }
        node = node.children[seg];
      }
      node.questionIds.push(q.id);
    }

    // 为所有节点计算累计题目数
    if (hasChapters) {
      _countTree(root);
      return root;
    }
    return null;
  }

  /** 从章节树中收集所有叶节点的问题 ID */
  static collectQuestionIds(chapterNode) {
    const ids = new Set(chapterNode.questionIds || []);
    if (chapterNode.children) {
      for (const child of Object.values(chapterNode.children)) {
        for (const id of ChapterTreeBuilder.collectQuestionIds(child)) {
          ids.add(id);
        }
      }
    }
    return [...ids];
  }
}

function _countTree(node) {
  let count = node.questionIds.length;
  for (const child of Object.values(node.children)) {
    count += _countTree(child);
  }
  node.count = count;
  return count;
}

/** 检测是否为安卓环境 (Mobile/Tablet) */
function isAndroid() {
  // Obsidian Android uses Capacitor, UA usually contains Android
  return /android/i.test(navigator.userAgent) ||
         (window.Capacitor && window.Capacitor.getPlatform() === 'android');
}

/** v2.7.6: 是否为作答中的模式（决定容器是否加 .quiz-question-view，见 CSS 第 33 节） */
function isQuestionMode(mode) {
  return mode === 'practice' || mode === 'review' || mode === 'wrong'
      || mode === 'favorites' || mode === 'group';
}

/* ============================================================
   Section 2: 表格解析器（支持 CSV/TSV，自动检测分隔符）
   ============================================================ */

class TableParser {
  static detectDelimiter(content) {
    const firstLineEnd = content.indexOf('\n');
    const headerLine = (firstLineEnd > 0 ? content.substring(0, firstLineEnd) : content)
      .replace(/\r$/, '');
    const stripped = headerLine.replace(/"[^"]*"/g, '');
    const tabCount = (stripped.match(/\t/g) || []).length;
    const commaCount = (stripped.match(/,/g) || []).length;
    return tabCount > commaCount ? '\t' : ',';
  }

  static parse(content, delimiter) {
    if (!delimiter) {
      delimiter = this.detectDelimiter(content);
    }
    const rows = [];
    let currentRow = [];
    let currentField = '';
    let inQuotes = false;
    let i = 0;

    while (i < content.length) {
      const char = content[i];
      if (inQuotes) {
        if (char === '"') {
          if (i + 1 < content.length && content[i + 1] === '"') {
            currentField += '"';
            i += 2;
          } else {
            inQuotes = false;
            i++;
          }
        } else {
          currentField += char;
          i++;
        }
      } else {
        if (char === '"' && currentField === '') {
          inQuotes = true;
          i++;
        } else if (char === '"') {
          currentField += char;
          i++;
        } else if (char === delimiter) {
          currentRow.push(currentField);
          currentField = '';
          i++;
        } else if (char === '\n' || char === '\r') {
          if (char === '\r' && i + 1 < content.length && content[i + 1] === '\n') {
            i++;
          }
          currentRow.push(currentField);
          currentField = '';
          if (currentRow.some(f => f.trim() !== '')) {
            rows.push(currentRow);
          }
          currentRow = [];
          i++;
        } else {
          currentField += char;
          i++;
        }
      }
    }
    if (currentField !== '' || currentRow.length > 0) {
      currentRow.push(currentField);
      if (currentRow.some(f => f.trim() !== '')) {
        rows.push(currentRow);
      }
    }
    return rows;
  }

  static parseOptions(optionsStr) {
    if (!optionsStr || isEmptyValue(optionsStr)) return [];
    const rawParts = optionsStr.split(/\||\n|\r/).map(p => p.trim()).filter(p => p !== '');
    let parts = rawParts;
    if (rawParts.length <= 1) {
      // v2.7.4: 判定条件改为「以字母标记开头 + 至少 2 个标记 + 字母严格递增」。
      // 原来要求 >= 3 个标记，于是 "A.对 B.错" 这种只有两个选项的单元格整段并成一个选项，
      // 答案字母找不到对应选项 → 该题永远判错，界面也只显示一个选项。
      // 加上"递增 + 开头"两个条件，则不会把 "I have a dream" 之类的英文句子切开。
      const markerPattern = /(?:^|\s+)([A-Za-z])\s*[-.、．)）：:）:]/g;
      const letters = [];
      let mk;
      while ((mk = markerPattern.exec(optionsStr)) !== null) letters.push(mk[1].toUpperCase());
      let markerMode = letters.length >= 2 && /^[A-Za-z]\s*[-.、．)）：:）:]/.test(optionsStr);
      for (let i = 1; markerMode && i < letters.length; i++) {
        if (letters[i].charCodeAt(0) <= letters[i - 1].charCodeAt(0)) markerMode = false;
      }
      if (markerMode) {
        parts = optionsStr.split(/\s+(?=[A-Za-z]\s*[-.、．)）：:）:])/)
          .map(p => p.trim()).filter(p => p !== '');
      }
    }
    const options = [];
    for (const part of parts) {
      // v2.6.2: 修复化学式/单词类选项首字母被误吞（如 NaCl|H₂O|Na₂CO₃|C₆H₁₂O₆）。
      // v2.7.4: 再把"单字母 + 空白"这一形式收紧 —— 只有当该字母正好等于这个位置
      // 应有的选项号时才认作选项号。否则 "I have a dream"、"A big apple" 这类
      // 英文选项文本会被吞掉首字母（解析成 key=I / text=have a dream）。
      // 带分隔符的形式（A. / A、 / A-）不受影响，仍按写明的字母取 key。
      const expected = String.fromCharCode(65 + options.length);
      const delim = part.match(/^([A-Za-z])\s*[-.、．)）：:）:]\s*(.+)$/);
      let key = null;
      let text = null;
      if (delim) {
        key = delim[1].toUpperCase();
        text = delim[2].trim();
      } else {
        const ws = part.match(/^([A-Za-z])\s+(.+)$/);
        if (ws && ws[1].toUpperCase() === expected) {
          key = expected;
          text = ws[2].trim();
        }
      }
      if (key && text) {
        options.push({ key, text });
      } else {
        options.push({ key: expected, text: part.trim() });
      }
    }
    return options;
  }

  static parseToQuestions(content) {
    if (content.codePointAt(0) === 0xFEFF) {
      content = content.slice(1);
    }
    const delimiter = this.detectDelimiter(content);
    const rows = this.parse(content, delimiter);
    if (rows.length < 2) {
      throw new Error('至少需要包含表头和一行数据');
    }
    const headers = rows[0].map(h => h.trim().replace(/[\u00A0\u2000-\u200F\u2028-\u202F\uFEFF]+/g, ''));
    const colMap = this.detectColumns(headers);

    // v2.6.2: 修复列索引为 0 时被 falsy 判断误判
    // （题干/答案列位于第一列时 colMap.question/answer === 0，原 !colMap.question 会误报"未找到列"）
    if (colMap.question === null) {
      throw new Error('未找到"题目/题干/question"列，请检查表头。当前表头：' + headers.join(', '));
    }
    if (colMap.answer === null) {
      throw new Error('未找到"答案/answer"列，请检查表头。当前表头：' + headers.join(', '));
    }

    const questions = [];
    const errors = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      try {
        const q = this.parseRow(row, colMap, headers, i + 1);
        // v2.7.4: 导入时就规范化一次，与"下次启动载入"保持一致，
        // 避免同一份数据在导入当次与重启后表现不同（选项 key 跳号等）。
        if (q) questions.push(normalizeQuestionData(q));
      } catch (e) {
        errors.push(`第 ${i + 1} 行: ${e.message}`);
      }
    }
    // v2.4: 构建章节树（仅当存在章节列时）
    const chapterTree = colMap.chapter !== null
      ? ChapterTreeBuilder.build(questions)
      : null;

    return { questions, errors, colMap, delimiter, chapterTree };
  }

  /**
   * 优化后的列检测逻辑
   * 1. 识别所有以"选项"开头的列
   * 2. 如果有多列，按顺序映射为 A, B, C...
   * 3. 如果只有一列，视为合并选项列
   */
  static detectColumns(headers) {
    const map = {
      question: null,
      options: null,       // 合并选项列
      optionCols: [],      // 分散选项列 [{key, idx, name}]
      answer: null,
      type: null,
      chapter: null,       // 章节列（v2.4 新增）
      extra: []
    };

    const questionExact = ['题干', 'question', '问题', '题目', '试题', '内容'];
    const answerExact = ['答案', 'answer', '正确答案', '参考答案'];
    const typeExact = ['题型', '类型', 'type', '题目类型'];

    // 辅助函数：检查是否以"选项"开头
    // v2.7.4: 修正正则优先级 —— 原 /^选项|option/i 实际含义是 (^选项) | (option)，
    // 任何位置含 option 的列都会被当成选项列。
    const isOptionCol = (h) => /^(?:选项|option)/i.test(h);
    // v2.7.4: 模糊匹配时排除"解析/答案/说明"类列，避免「题目解析」被当成题干列静默导入
    const looksLikeMeta = (h) => /解析|答案|说明|备注|依据|来源|分值|难度/.test(h);

    // 第一遍：精确匹配非选项列
    headers.forEach((h, idx) => {
      const lower = h.toLowerCase().trim();
      if (map.question === null && questionExact.includes(lower)) map.question = idx;
      else if (map.answer === null && answerExact.includes(lower)) map.answer = idx;
      else if (map.type === null && typeExact.includes(lower)) map.type = idx;
    });

    // 第二遍：模糊匹配非选项列（跳过"解析/答案/说明"类列）
    headers.forEach((h, idx) => {
      const lower = h.toLowerCase().trim();
      if (map.question === null && !looksLikeMeta(h) && ['题干', 'question', '问题', '题目'].some(k => lower.includes(k))) map.question = idx;
      else if (map.answer === null && !/解析|说明|备注/.test(h) && ['答案', 'answer'].some(k => lower.includes(k))) map.answer = idx;
      else if (map.type === null && ['题型', '类型', 'type'].some(k => lower.includes(k))) map.type = idx;
    });

    // 第三遍：识别章节列（v2.4 新增：支持 > 分隔的层级路径）
    const chapterExact = ['章节', 'chapter', '章', '节', 'section', 'unit'];
    headers.forEach((h, idx) => {
      if (map.chapter === null && (chapterExact.includes(h.toLowerCase().trim()) || 
          ['章节', 'chapter', '章', '节'].some(k => h.toLowerCase().includes(k)))) {
        // 确保不是已分配的列
        if (map.question !== idx && map.answer !== idx && map.type !== idx) {
          map.chapter = idx;
        }
      }
    });

    // 第四遍：处理选项列
    // 收集所有疑似选项列
    const potentialOptionCols = [];
    headers.forEach((h, idx) => {
      // 跳过已分配的列（包括章节列）
      if (map.question === idx || map.answer === idx || map.type === idx || map.chapter === idx) return;

      if (isOptionCol(h)) {
        potentialOptionCols.push({
          originalName: h,
          idx
        });
      } else {
        map.extra.push({ name: h, idx });
      }
    });

    if (potentialOptionCols.length > 0) {
      // 情况 A: 多列选项 (例如: 选项A, 选项B, 选项C)
      if (potentialOptionCols.length > 1) {
        // 按原始顺序或检测到的字母排序？通常按列顺序更稳妥，除非字母明确
        // 这里我们按列索引顺序分配 A, B, C...
        potentialOptionCols.sort((a, b) => a.idx - b.idx);
        potentialOptionCols.forEach((col, i) => {
          map.optionCols.push({
            key: String.fromCharCode(65 + i), // A, B, C...
            idx: col.idx,
            name: col.originalName
          });
        });
      }
      // 情况 B: 单列选项 (例如: 选项, Options) -> 走原有合并解析逻辑
      else {
        map.options = potentialOptionCols[0].idx;
      }
    }

    return map;
  }

  static parseRow(row, colMap, headers, lineNum) {
    // v2.7.3: 一并考虑章节列，避免章节位于末列时读取越界
    const minCols = Math.max(
      colMap.question,
      colMap.answer,
      colMap.type !== null ? colMap.type : 0,
      colMap.chapter !== null ? colMap.chapter : 0,
      ...colMap.optionCols.map(o => o.idx),
      ...colMap.extra.map(e => e.idx)
    );
    while (row.length <= minCols) {
      row.push('');
    }

    const text = (row[colMap.question] || '').trim().replace(/\t/g, ' ').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    if (!text || isEmptyValue(text)) return null;

    const answerRaw = (row[colMap.answer] || '').trim();
    if (!answerRaw || isEmptyValue(answerRaw)) return null;

    let answerClean = answerRaw.replace(/[\[\]\(\)\{\}""''\s,，;；]+/g, '').toUpperCase();
    if (!answerClean || isEmptyValue(answerClean)) return null;

    let options = [];
    if (colMap.options !== null) {
      options = this.parseOptions(row[colMap.options] || '');
    } else if (colMap.optionCols.length > 0) {
      for (const opt of colMap.optionCols) {
        const val = (row[opt.idx] || '').trim().replace(/\t/g, ' ').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        // v2.7.4: 只跳过真正的空单元格。原先用 isEmptyValue 会把「无 / - / N/A / \」这类
        // 合法选项内容一起丢掉，导致选项字母跳号、答案与选项错位而判分错误。
        if (val !== '') {
          options.push({ key: opt.key, text: val });
        }
      }
    }

    let type = null;
    if (colMap.type !== null && row[colMap.type]) {
      const t = (row[colMap.type] || '').trim().toLowerCase();
      if (t.includes('多选') || t.includes('multiple')) type = QUESTION_TYPES.MULTIPLE;
      else if (t.includes('判断') || t.includes('judge')) type = QUESTION_TYPES.JUDGE;
      else if (t.includes('单选') || t.includes('single')) type = QUESTION_TYPES.SINGLE;
    }
    if (!type) {
      type = this.inferType(answerClean, options);
    }

    if (type === QUESTION_TYPES.JUDGE && options.length === 0) {
      options = [
        { key: 'A', text: '正确' },
        { key: 'B', text: '错误' }
      ];
    }

    let answer = answerClean;
    if (type === QUESTION_TYPES.JUDGE) {
      // v2.7.4: 补充 Y/N、YES/NO、1/0 等常见判断题写法
      if (['对', '√', 'T', 'TRUE', '正确', '是', 'Y', 'YES', '1'].includes(answerClean)) answer = 'A';
      else if (['错', '×', 'F', 'FALSE', '错误', '否', 'N', 'NO', '0'].includes(answerClean)) answer = 'B';
    } else {
      // v2.6.1: 选择题答案规范化（全角→半角、大写、剔除「、」等杂字符）
      answer = normalizeAnswerLetters(answerClean) || answerClean;
    }

    const extraFields = {};
    for (const extra of colMap.extra) {
      const val = (row[extra.idx] || '').trim().replace(/\t/g, ' ').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      if (!isEmptyValue(val)) {
        extraFields[extra.name] = val;
      }
    }

    // v2.7.3: 标签列同时保留「标签筛选」与「解析显示」两种用途。
    // 原先提取后会从 extraFields 删除，导致解析里看不到标签；现在保留。
    const tags = [];
    if (extraFields['标签'] || extraFields['tags'] || extraFields['tag']) {
      const tagStr = extraFields['标签'] || extraFields['tags'] || extraFields['tag'];
      tags.push(...tagStr.split(/[,，;；|]/).map(t => t.trim()).filter(Boolean));
      // 不再 delete：标签内容一并显示在解析中
    }

    // v2.7.3: 章节列不属于核心列，除构建章节树外，内容也放进解析
    if (colMap.chapter !== null) {
      const chapterName = headers[colMap.chapter] || '章节';
      const chapterVal = (row[colMap.chapter] || '').trim().replace(/\t/g, ' ');
      if (!isEmptyValue(chapterVal) && !(chapterName in extraFields)) {
        extraFields[chapterName] = chapterVal;
      }
    }

    // v2.7.4: 「解析 / 答案解析 / 题目解析…」列归位到 analysis（与编辑弹窗同一个字段）
    const analysisText = extractAnalysisField(extraFields);

    // v2.4: 提取章节路径（兼容 > 分隔的层级章节）
    let chapterPath = null;
    if (colMap.chapter !== null) {
      const chapterRaw = (row[colMap.chapter] || '').trim();
      if (chapterRaw && !isEmptyValue(chapterRaw)) {
        chapterPath = chapterRaw
          .split(/>/)
          .map(s => s.trim())
          .filter(s => s.length > 0);
        if (chapterPath.length === 0) chapterPath = null;
      }
    }

    return {
      id: generateId(),
      text,
      type,
      options,
      answer,
      analysis: analysisText,
      tags,
      extraFields,
      chapterPath,       // v2.4: 章节层级路径 ['L1','L2','L3'] 或 null
      lineNum
    };
  }

  static inferType(answerRaw, options) {
    if (answerRaw.length > 1 && /^[A-Z]+$/.test(answerRaw)) return QUESTION_TYPES.MULTIPLE;
    // v2.7.4: 加入 Y/N、YES/NO；刻意不加 '1'/'0' —— 有些题库用 1/2/3/4 表示选项序号，
    // 一旦被判成判断题，整题选项都会被替换成"正确/错误"，损失远大于收益。
    const judgeAnswers = ['对', '错', '√', '×', 'T', 'F', 'TRUE', 'FALSE', '正确', '错误', '是', '否', 'Y', 'N', 'YES', 'NO'];
    if (judgeAnswers.includes(answerRaw)) return QUESTION_TYPES.JUDGE;
    if (options.length === 2) {
      const optTexts = options.map(o => o.text.toLowerCase());
      if (optTexts.some(t => t.includes('正确') || t.includes('对')) &&
          optTexts.some(t => t.includes('错误') || t.includes('错'))) {
        return QUESTION_TYPES.JUDGE;
      }
    }
    return QUESTION_TYPES.SINGLE;
  }
}

/* ============================================================
   Section 3: 相似度服务
   ============================================================ */

class SimilarityService {
  static diceCoefficient(str1, str2) {
    if (!str1 || !str2) return 0;
    const s1 = str1.toLowerCase().replace(/\s+/g, '');
    const s2 = str2.toLowerCase().replace(/\s+/g, '');
    if (s1 === s2) return 1;
    if (s1.length < 2 || s2.length < 2) return 0;
    const bigrams1 = new Map();
    for (let i = 0; i < s1.length - 1; i++) {
      const bg = s1.substring(i, i + 2);
      bigrams1.set(bg, (bigrams1.get(bg) || 0) + 1);
    }
    let intersection = 0;
    for (let i = 0; i < s2.length - 1; i++) {
      const bg = s2.substring(i, i + 2);
      const count = bigrams1.get(bg);
      if (count > 0) {
        intersection++;
        bigrams1.set(bg, count - 1);
      }
    }
    return (2 * intersection) / (s1.length - 1 + s2.length - 1);
  }

  static calculateScore(q1, q2) {
    if (q1.type !== q2.type) return 0;
    const textScore = this.diceCoefficient(q1.text, q2.text);
    let tagScore = 0;
    if (q1.tags.length > 0 && q2.tags.length > 0) {
      const common = q1.tags.filter(t => q2.tags.includes(t)).length;
      tagScore = common / Math.max(q1.tags.length, q2.tags.length);
    }
    let optScore = 0;
    if (q1.options.length > 0 && q2.options.length > 0) {
      const minLen = Math.min(q1.options.length, q2.options.length);
      let sum = 0;
      for (let i = 0; i < minLen; i++) {
        sum += this.diceCoefficient(q1.options[i].text, q2.options[i].text);
      }
      optScore = sum / minLen;
    }
    return textScore * 0.5 + tagScore * 0.2 + optScore * 0.3;
  }

  static findSimilar(target, allQuestions, threshold = 0.3, maxResults = 5) {
    const candidates = allQuestions.filter(q => q.id !== target.id && q.type === target.type);
    const scored = candidates.map(q => ({ question: q, score: this.calculateScore(target, q) }));
    return scored.filter(s => s.score >= threshold).sort((a, b) => b.score - a.score).slice(0, maxResults);
  }
}

/* ============================================================
   Section 4: 确认对话框
   ============================================================ */

class ConfirmModal extends Modal {
  constructor(app, title, message, onConfirm, confirmText = '确定', cancelText = '取消') {
    super(app);
    this._title = title;
    this._message = message;
    this._onConfirm = onConfirm;
    this._confirmText = confirmText;
    this._cancelText = cancelText;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('div', { cls: 'quiz-confirm-title', text: this._title });
    contentEl.createEl('div', { cls: 'quiz-confirm-message', text: this._message });
    const actions = contentEl.createEl('div', { cls: 'quiz-confirm-actions' });
    const cancelBtn = actions.createEl('button', { cls: 'quiz-btn', text: this._cancelText });
    cancelBtn.addEventListener('click', () => this.close());
    const confirmBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-danger', text: this._confirmText });
    confirmBtn.addEventListener('click', () => { this._onConfirm(); this.close(); });
  }
  onClose() { this.contentEl.empty(); }
}

function showConfirm(app, title, message, confirmText = '确定') {
  return new Promise(resolve => {
    let confirmed = false;
    const modal = new ConfirmModal(app, title, message, () => { confirmed = true; resolve(true); }, confirmText);
    const origClose = modal.onClose.bind(modal);
    modal.onClose = () => { origClose(); if (!confirmed) resolve(false); };
    modal.open();
  });
}

/* ============================================================
   Section 5: 题目列表模态框基类
   ============================================================ */

class QuestionListModal extends Modal {
  constructor(app, plugin, view) {
    super(app);
    this.plugin = plugin;
    this.view = view;
  }
  getTitle() { return '题目列表'; }
  getQuestionIds() { return []; }
  getEmptyText() { return '暂无数据'; }
  getPracticeLabel() { return '开始练习'; }
  async onRemoveItem(questionId) { }
  getBank() { return this.view ? this.view.currentBank : null; }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    this.renderList(contentEl);
  }

  renderList(container) {
    container.empty();
    container.createEl('h3', { cls: 'quiz-modal-title', text: this.getTitle() });
    const bank = this.getBank();
    if (!bank) {
      container.createEl('div', { cls: 'quiz-empty-text', text: '请先选择题库' });
      return;
    }
    const ids = this.getQuestionIds();
    const idSet = new Set(ids);   // v2.7.4: 万题级题库下 O(n·m) 会明显卡顿
    const questions = bank.questions.filter(q => idSet.has(q.id));
    if (questions.length === 0) {
      const empty = container.createEl('div', { cls: 'quiz-empty' });
      empty.createEl('div', { cls: 'quiz-empty-icon', text: '📭' });
      empty.createEl('div', { cls: 'quiz-empty-text', text: this.getEmptyText() });
      return;
    }
    const practiceBtn = container.createEl('button', {
      cls: 'quiz-btn quiz-btn-primary quiz-list-practice-btn',
      text: `${this.getPracticeLabel()}（${questions.length} 题）`
    });
    practiceBtn.addEventListener('click', () => { this.close(); this.startPractice(questions); });
    const listEl = container.createEl('div');
    questions.forEach(q => this.renderQuestionItem(listEl, q));
  }

  renderQuestionItem(container, question) {
    const item = container.createEl('div', { cls: 'quiz-list-item' });
    const header = item.createEl('div', { cls: 'quiz-list-item-header' });
    header.createEl('span', { cls: 'quiz-question-type', text: TYPE_LABELS[question.type] || question.type });
    const textEl = header.createEl('div', { cls: 'quiz-list-item-text' });
    safeSetText(textEl, question.text);
    const actions = header.createEl('div', { cls: 'quiz-list-item-actions' });
    // 复制题目为 Callout（优先走 view 方法以自动定位原题；view 缺失时直接调用工具函数）
    const copyBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: '📋 复制' });
    copyBtn.title = '复制题目为 Callout 格式';
    copyBtn.setAttribute('aria-label', '复制题目');
    copyBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const bank = this.getBank();
      if (this.view && typeof this.view.copyQuestionToClipboard === 'function') {
        this.view.copyQuestionToClipboard(question, bank ? bank.name : '');
      } else {
        copyQuestionToClipboard(question, bank ? bank.name : '');
      }
    });
    const removeBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-danger', text: '移除' });
    removeBtn.addEventListener('click', async (e) => {
      e.stopPropagation();
      await this.onRemoveItem(question.id);
      this.renderList(this.contentEl);
    });
    if (question.options && question.options.length > 0) {
      const optEl = item.createEl('div', { cls: 'quiz-list-item-options' });
      question.options.forEach(opt => {
        const line = optEl.createEl('div');
        line.createEl('strong', { text: `${opt.key}. ` });
        line.createSpan({ text: opt.text });
      });
    }
    const ansEl = item.createEl('div', { cls: 'quiz-list-item-answer' });
    safeSetText(ansEl, `答案：${question.answer}`);
    this.renderExtraFields(item, question);
  }

  renderExtraFields(container, question) {
    if (!question.extraFields || Object.keys(question.extraFields).length === 0) return;
    const extraEl = container.createEl('div', { cls: 'quiz-analysis' });
    Object.entries(question.extraFields).forEach(([key, val]) => {
      if (isEmptyValue(val)) return;
      const field = extraEl.createEl('div', { cls: 'quiz-extra-field' });
      field.createEl('div', { cls: 'quiz-extra-field-title', text: key });
      field.createEl('div', { cls: 'quiz-extra-field-content', text: val });
    });
  }

  startPractice(questions) {
    if (this.view) this.view.startPracticeMode(questions, this.getPracticeLabel());
  }
  onClose() { this.contentEl.empty(); }
}

/* ============================================================
   Section 6-10: 各类模态框 (Wrong, Favorites, AnswerSheet, Similar, Edit)
   ============================================================ */

class WrongQuestionsModal extends QuestionListModal {
  getTitle() { return '📕 错题集'; }
  getEmptyText() { return '太棒了，没有错题！'; }
  getPracticeLabel() { return '错题重练'; }
  getQuestionIds() {
    const bank = this.getBank();
    if (!bank) return [];
    return this.plugin.getWrongQuestions(bank.id);
  }
  async onRemoveItem(questionId) {
    const bank = this.getBank();
    if (bank) {
      this.plugin.removeWrongQuestion(bank.id, questionId);
      await this.plugin.saveDataDebounced();
      new Notice('已从错题集移除');
    }
  }
  /** v2.7.4: "清空错题集"并入列表渲染。
   *  原来只在 onOpen 里追加，移除任意一条错题后 renderList 重建 contentEl，按钮就消失了。 */
  renderList(container) {
    super.renderList(container);
    const bank = this.getBank();
    if (!bank) return;
    if (this.getQuestionIds().length === 0) return;
    const clearBtn = container.createEl('button', { cls: 'quiz-btn quiz-btn-danger quiz-btn-sm quiz-list-clear-btn', text: '清空错题集' });
    clearBtn.addEventListener('click', async () => {
      const confirmed = await showConfirm(this.app, '清空错题集', '确定要清空所有错题吗？此操作不可撤销。');
      if (confirmed) {
        this.plugin.clearWrongQuestions(bank.id);
        await this.plugin.saveDataDebounced();
        this.renderList(this.contentEl);
        new Notice('错题集已清空');
      }
    });
  }
}

class FavoritesModal extends QuestionListModal {
  getTitle() { return '⭐ 收藏题目'; }
  getEmptyText() { return '还没有收藏任何题目'; }
  getPracticeLabel() { return '收藏练习'; }
  getQuestionIds() {
    const bank = this.getBank();
    if (!bank) return [];
    return this.plugin.getFavoriteQuestions(bank.id);
  }
  async onRemoveItem(questionId) {
    const bank = this.getBank();
    if (bank) {
      this.plugin.removeFavoriteQuestion(bank.id, questionId);
      await this.plugin.saveDataDebounced();
      new Notice('已取消收藏');
    }
  }
}

class AnswerSheetModal extends Modal {
  constructor(app, view) { super(app); this.view = view; }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', { cls: 'quiz-modal-title', text: '📋 答题卡' });
    const questions = this.view.currentQuestions;
    if (!questions || questions.length === 0) {
      contentEl.createEl('div', { cls: 'quiz-empty-text', text: '暂无题目' });
      return;
    }
    const groups = {};
    questions.forEach((q, idx) => {
      const type = q.type || 'other';
      if (!groups[type]) groups[type] = [];
      groups[type].push({ question: q, index: idx });
    });
    const records = this.view.answerRecords || {};
    Object.entries(groups).forEach(([type, items]) => {
      const section = contentEl.createEl('div', { cls: 'quiz-modal-section' });
      section.createEl('div', { cls: 'quiz-modal-section-title', text: `${TYPE_LABELS[type] || type}（${items.length} 题）` });
      const grid = section.createEl('div', { cls: 'quiz-sheet-grid' });
      items.forEach(({ question, index }, seqInGroup) => {
        const cell = grid.createEl('div', { cls: 'quiz-sheet-cell', text: String(seqInGroup + 1) });
        cell.title = `第 ${index + 1} 题`;
        const record = records[question.id];
        if (record) cell.addClass(record.correct ? 'answered' : 'wrong');
        if (index === this.view.currentIndex) cell.addClass('current');
        cell.addEventListener('click', () => { this.view.goToQuestion(index); this.close(); });
      });
    });
    const total = questions.length;
    // v2.7.4: 只统计"当前答题卡包含的题目"。原来直接数整个模式记录，会掺入历史记录：
    // 例如错题集清空后重练 10 题，模式记录里还留着旧记录，会出现"已作答 100 / 未作答 -90"。
    const sheetQids = new Set(questions.map(q => q.id));
    const sheetRecords = Object.entries(records).filter(([qid]) => sheetQids.has(qid));
    const answered = sheetRecords.length;
    const wrongCount = sheetRecords.filter(([, r]) => r && !r.correct).length;
    const stats = contentEl.createEl('div', { cls: 'quiz-stats quiz-sheet-stats' });
    [
      { value: total, label: '总题数' },
      { value: answered, label: '已作答' },
      { value: total - answered, label: '未作答' },
      { value: wrongCount, label: '答错' }
    ].forEach(s => {
      const card = stats.createEl('div', { cls: 'quiz-stat-card' });
      card.createEl('div', { cls: 'quiz-stat-value', text: String(s.value) });
      card.createEl('div', { cls: 'quiz-stat-label', text: s.label });
    });
    // v2.4.1: 清除当前答题卡的答题记录（不影响错题集与收藏）
    const clearBtn = contentEl.createEl('button', { cls: 'quiz-btn quiz-btn-danger quiz-btn-sm quiz-sheet-clear-btn', text: '清除答题记录' });
    clearBtn.addEventListener('click', async () => {
      const confirmed = await showConfirm(this.app, '清除答题记录', '确定要清除当前答题卡的答题记录吗？\n错题记录与收藏记录不会受到影响。');
      if (!confirmed) return;
      const bank = this.view.currentBank;
      const qids = questions.map(q => q.id);
      // 1. 清空当前答题卡的内存记录
      this.view.answerRecords = {};
      // 2. 清空持久化记录（仅当前答题卡涉及的题目）：当前模式记录 + 全局记录
      if (bank) {
        const practiceKey = this.view.getPracticeKey();
        const modeRecs = this.view.plugin.getModeAnswerRecords(bank.id, practiceKey);
        qids.forEach(id => { delete modeRecs[id]; });
        qids.forEach(id => { this.view.plugin.removeAnswerRecord(bank.id, id); });
        await this.view.plugin.saveDataDebounced();
      }
      // 3. 重置主视图当前题状态并刷新（保留当前所在题，仅清掉作答状态）
      if (this.view.currentQuestions.length > 0) {
        this.view.resetQuestionState();
        this.view.render();
      }
      new Notice('答题记录已清除');
      this.close();
    });
  }
  onClose() { this.contentEl.empty(); }
}

class SimilarQuestionsModal extends Modal {
  constructor(app, plugin, view, targetQuestion) {
    super(app);
    this.plugin = plugin;
    this.view = view;
    this.targetQuestion = targetQuestion;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', { cls: 'quiz-modal-title', text: '🔍 相似题推荐' });
    const bank = this.view.currentBank;
    if (!bank) return;
    // 使用设置中的相似度阈值（原来硬编码 0.25，导致设置项形同虚设）
    const threshold = typeof this.plugin.settings.similarityThreshold === 'number'
      ? this.plugin.settings.similarityThreshold
      : 0.3;
    const similar = SimilarityService.findSimilar(this.targetQuestion, bank.questions, threshold, 8);
    if (similar.length === 0) {
      contentEl.createEl('div', { cls: 'quiz-empty-text', text: '未找到相似题目' });
      return;
    }
    similar.forEach(({ question, score }) => {
      const item = contentEl.createEl('div', { cls: 'quiz-similar-item' });
      item.createEl('span', { cls: 'quiz-similar-score', text: `相似度 ${(score * 100).toFixed(0)}%` });
      const textEl = item.createEl('div', { cls: 'quiz-similar-text', text: question.text });
      item.createEl('div', { text: `答案：${question.answer}`, cls: 'quiz-list-item-answer' });
      item.addEventListener('click', () => {
        this.close();
        const idx = this.view.currentQuestions.findIndex(q => q.id === question.id);
        if (idx !== -1) { this.view.goToQuestion(idx); }
        else { new Notice('该题不在当前练习中，请从题库列表进入'); }
      });
    });
  }
  onClose() { this.contentEl.empty(); }
}

class EditQuestionModal extends Modal {
  constructor(app, plugin, view, question) {
    super(app);
    this.plugin = plugin;
    this.view = view;
    this.question = JSON.parse(JSON.stringify(question));
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', { cls: 'quiz-modal-title', text: '✏️ 编辑题目' });
    const typeField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    typeField.createEl('label', { cls: 'quiz-edit-label', text: '题型' });
    const typeSelect = typeField.createEl('select', { cls: 'quiz-edit-input' });
    Object.entries(TYPE_LABELS).forEach(([val, label]) => {
      const opt = typeSelect.createEl('option', { value: val, text: label });
      if (val === this.question.type) opt.selected = true;
    });
    const textField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    textField.createEl('label', { cls: 'quiz-edit-label', text: '题干' });
    const textInput = textField.createEl('textarea', { cls: 'quiz-edit-textarea', text: this.question.text });
    const optionsField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    optionsField.createEl('label', { cls: 'quiz-edit-label', text: '选项（每行一个，格式：A. 选项内容）' });
    const optionsInput = optionsField.createEl('textarea', {
      cls: 'quiz-edit-textarea',
      text: this.question.options.map(o => `${o.key}. ${o.text}`).join('\n')
    });
    const ansField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    ansField.createEl('label', { cls: 'quiz-edit-label', text: '答案（如：A 或 ABC）' });
    const ansInput = ansField.createEl('input', { cls: 'quiz-edit-input', value: this.question.answer });
    const anaField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    anaField.createEl('label', { cls: 'quiz-edit-label', text: '解析' });
    const anaInput = anaField.createEl('textarea', { cls: 'quiz-edit-textarea', text: this.question.analysis || '' });
    const tagField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    tagField.createEl('label', { cls: 'quiz-edit-label', text: '标签（逗号分隔）' });
    const tagInput = tagField.createEl('input', { cls: 'quiz-edit-input', value: (this.question.tags || []).join(', ') });
    const errorEl = contentEl.createEl('div', { cls: 'quiz-edit-error' });
    errorEl.hidden = true;
    const saveBtn = contentEl.createEl('button', { cls: 'quiz-btn quiz-btn-primary quiz-edit-save-btn', text: '保存修改' });
    saveBtn.addEventListener('click', async () => {
      const errors = this.validate(typeSelect.value, textInput.value, optionsInput.value, ansInput.value);
      if (errors.length > 0) {
        errorEl.hidden = false;
        errorEl.textContent = errors.join('；');
        return;
      }
      errorEl.hidden = true;
      const newOptions = optionsInput.value.split('\n').map(l => l.trim()).filter(Boolean).map(line => {
        const match = line.match(/^([A-Z])\s*[.、．\-)）]\s*(.+)$/i);
        return match ? { key: match[1].toUpperCase(), text: match[2].trim() } : null;
      }).filter(Boolean);
      this.question.type = typeSelect.value;
      this.question.text = textInput.value.trim();
      this.question.options = newOptions;
      this.question.answer = normalizeAnswerLetters(ansInput.value) || ansInput.value.trim().toUpperCase();
      this.question.analysis = anaInput.value.trim();
      this.question.tags = tagInput.value.split(/[,，]/).map(t => t.trim()).filter(Boolean);
      await this.plugin.updateQuestion(this.view.currentBank.id, this.question);
      await this.plugin.saveDataDebounced();
      // 同步更新当前练习列表中的对应题（避免编辑后仍显示旧内容）
      const viewIdx = this.view.currentQuestions.findIndex(x => x.id === this.question.id);
      if (viewIdx !== -1) this.view.currentQuestions[viewIdx] = this.question;
      this.view.render();
      new Notice('题目已更新');
      this.close();
    });
  }
  validate(type, text, optionsRaw, answer) {
    const errors = [];
    if (!text.trim()) errors.push('题干不能为空');
    const optionLines = optionsRaw.split('\n').map(l => l.trim()).filter(Boolean);
    const validOptions = optionLines.filter(l => /^[A-Z]\s*[.、．\-)）]\s*.+$/i.test(l));
    if (type !== QUESTION_TYPES.JUDGE && validOptions.length < 2) {
      errors.push('至少需要2个有效选项（格式：A. 内容）');
    }
    if (!answer.trim()) errors.push('答案不能为空');
    const optionKeys = validOptions.map(l => l.match(/^([A-Z])/i)[1].toUpperCase());
    const answerKeys = answer.trim().toUpperCase().split('');
    if (type === QUESTION_TYPES.SINGLE && answerKeys.length !== 1) {
      errors.push('单选题答案只能有一个字母');
    }
    for (const key of answerKeys) {
      if (optionKeys.length > 0 && !optionKeys.includes(key)) {
        errors.push(`答案 "${key}" 不在选项范围内`);
      }
    }
    return errors;
  }
  onClose() { this.contentEl.empty(); }
}

/* ============================================================
   Section 10.5: 导入查重处理方式选择（v2.5）
   ============================================================ */

class DuplicateImportModal extends Modal {
  /**
   * 导入题目到已有题库时，若发现重复题（题干相同），让用户选择处理方式。
   * @param {number} dupCount 重复题数量
   * @param {number} newCount 新题数量
   * @param {Function} onChoose 回调：'copy' | 'replace' | 'ignore' | null（取消）
   */
  constructor(app, dupCount, newCount, onChoose) {
    super(app);
    this.dupCount = dupCount;
    this.newCount = newCount;
    this.onChoose = onChoose;
    this._chosen = false;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', { cls: 'quiz-modal-title', text: '⚠️ 发现重复题目' });
    const msg = contentEl.createEl('div', { cls: 'quiz-confirm-message' });
    msg.createEl('div', { text: `检测到 ${this.dupCount} 道题与当前题库中的题目题干相同，另有 ${this.newCount} 道新题。` });
    msg.createEl('div', { text: '新题将全部导入，请选择重复题目的处理方式：' });
    const actions = contentEl.createEl('div', { cls: 'quiz-dup-actions' });
    const mkBtn = (text, cls, action, desc) => {
      const b = actions.createEl('button', { cls: `quiz-btn ${cls}`.trim(), text });
      b.title = desc;
      b.addEventListener('click', () => { this._chosen = true; this.close(); this.onChoose(action); });
    };
    mkBtn('📄 复制导入 — 重复题作为新题加入（两份并存）', 'quiz-btn-success', 'copy', '重复题目也一并导入，题库中保留两道相同题干的题');
    mkBtn('🔄 替换导入 — 用导入内容覆盖题库中的原题', 'quiz-btn-primary', 'replace', '原题内容更新为导入版本，答题记录/错题/收藏关联保持不变');
    mkBtn('⏭ 忽略导入 — 跳过重复题，仅导入新题', '', 'ignore', '重复题目不导入');
    const cancelRow = contentEl.createEl('div', { cls: 'quiz-confirm-actions' });
    const cancelBtn = cancelRow.createEl('button', { cls: 'quiz-btn', text: '取消整个导入' });
    cancelBtn.addEventListener('click', () => { this._chosen = true; this.close(); this.onChoose(null); });
  }
  onClose() {
    if (!this._chosen) this.onChoose(null);
    this.contentEl.empty();
  }
}

/* ============================================================
   Section 11: CSV/TSV 导入模态框
   ============================================================ */

class ImportModal extends Modal {
  /**
   * @param targetBank 传入题库对象 = 导入题目到已有题库（v2.5）；为空 = 新建题库导入
   */
  constructor(app, plugin, view, targetBank = null) {
    super(app);
    this.plugin = plugin;
    this.view = view;
    this.targetBank = targetBank || null;
    this._fileContent = null;
    this._fileName = null;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', {
      cls: 'quiz-modal-title',
      text: this.targetBank ? `📥 导入题目到「${this.targetBank.name}」` : '📥 导入题库'
    });

    // 题库名称（导入到已有题库时无需填写）
    let nameInput = null;
    if (!this.targetBank) {
      const nameField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
      nameField.createEl('label', { cls: 'quiz-edit-label', text: '题库名称' });
      nameInput = nameField.createEl('input', {
        cls: 'quiz-edit-input',
        attr: { placeholder: '例如：培训考试' }
      });

      // 如果有文件名，自动填充
      if (this._fileName) {
        const baseName = this._fileName.replace(/\.[^/.]+$/, "");
        nameInput.value = baseName;
      }
    }

    // 拖拽区域
    const dropzone = contentEl.createEl('div', { cls: 'quiz-import-dropzone' });
    dropzone.createEl('div', { cls: 'quiz-import-dropzone-icon', text: '📄' });
    dropzone.createEl('div', { cls: 'quiz-import-dropzone-text', text: '拖拽文件到此处，或点击选择文件' });
    dropzone.createEl('div', { cls: 'quiz-import-dropzone-hint', text: '支持 .csv / .tsv / .txt / .xls / .xlsx 文件' });

    // 文件信息
    const fileInfo = contentEl.createEl('div', { cls: 'quiz-import-file-info' });
    fileInfo.hidden = true;

    // 手动粘贴区域
    const pasteField = contentEl.createEl('div', { cls: 'quiz-edit-field' });
    pasteField.createEl('label', { cls: 'quiz-edit-label', text: '或直接粘贴内容' });
    const csvArea = pasteField.createEl('textarea', {
      cls: 'quiz-import-area',
      attr: {
        placeholder: '题型\t题干\t选项\t答案\t解析\n单选题\t1+1等于?\tA-1|B-2|C-3|D-4\tB\t基础运算'
      }
    });

    // 格式说明（v2.7.4：改为「常显要点 + 可展开完整说明」，内容逐条对照解析代码核过）
    const formatInfo = contentEl.createEl('div', { cls: 'quiz-analysis' });
    formatInfo.createEl('div', { cls: 'quiz-analysis-title', text: '📋 导入格式说明' });

    /** 说明正文行 */
    const fmtLine = (parent, text) => parent.createEl('div', { cls: 'quiz-import-fmt-line', text });
    /** 分组小标题（内联加粗，避免改动样式表） */
    const fmtHead = (parent, text) => {
      const el = parent.createEl('div', { cls: 'quiz-import-fmt-line', text });
      el.style.fontWeight = '650';
      el.style.marginTop = '10px';
    };

    // —— 常显要点：最常用、最容易踩坑的 5 条 ——
    [
      '• 必须有三列：题干列 + 选项列 + 答案列。列顺序随意，其余列会自动收进"解析"区的附加信息',
      '  —— 判断题可以不给选项列，会自动补「A. 正确 / B. 错误」',
      '• 建议表头：题型 | 题干 | 选项 | 答案 | 解析 | 章节 | 标签',
      '• 答案写选项字母：单选 B；多选 ABD；判断 对 / 错（也认 √ × T F Y N 1 0、正确 / 错误）',
      '• 选项两种写法：合并一列「A-甲|B-乙|C-丙」，或分散多列「选项A / 选项B / 选项C…」，项数不限',
      '• 题干或答案为空的行会被整行跳过；导入前先点「🔍 预览解析结果」核对'
    ].forEach(line => fmtLine(formatInfo, line));

    // v2.7.4: 模板下载（模板随插件分发，位置：插件目录/题库导入模板.xlsx）
    const tplRow = formatInfo.createEl('div');
    tplRow.style.marginTop = '10px';
    tplRow.style.display = 'flex';
    tplRow.style.alignItems = 'center';
    tplRow.style.flexWrap = 'wrap';
    tplRow.style.gap = '8px';
    const tplBtn = tplRow.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-primary', text: '⬇️ 下载 Excel 模板' });
    tplBtn.title = '保存到本地：含 12 行示例数据、使用说明、写法对照、合并选项列示例';
    tplBtn.setAttribute('aria-label', '下载 Excel 模板');
    tplBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.downloadTemplate();
    });
    tplRow.createEl('span', { cls: 'quiz-import-fmt-line', text: '下载后按模板填写即可：4 个工作表（题库模板 / 使用说明 / 写法对照 / 合并选项列示例）' });

    const fmtDetail = formatInfo.createEl('details');
    fmtDetail.createEl('summary', { text: '展开：完整列名规则 / 选项写法 / 答案与题型 / 文件与编码 / 已导入题库与数据安全' });
    const fmtBox = fmtDetail.createEl('div');
    fmtBox.style.marginTop = '6px';

    fmtHead(fmtBox, '① 列名怎么认');
    [
      '• 题干列（必须）：题干 / 题目 / 问题 / 试题 / 内容 / question；名称里含「题干、题目、问题、question」的列也能识别',
      '• 答案列（必须）：答案 / 正确答案 / 参考答案 / answer；名称里含「答案、answer」的列也能识别',
      '• 题型列（可选）：题型 / 类型 / 题目类型 / type',
      '• 选项列（必须）：表头必须以「选项」或「option」开头，大小写不敏感。能识别的写法例如：选项、选项A、选项B、选项C、选项1、选项一、options、OptionA',
      '  —— 只有一列时按②的合并写法解析；有多列时按列的先后顺序编号 A、B、C…（列名里写的字母不参与编号）',
      '  —— 单选题 / 多选题没有选项列就无法作答（该题会被判错）；判断题可以不给选项列，会自动补「A. 正确 / B. 错误」',
      '  —— 不被识别的常见写法：表头写成 A / B / C / D、选择项、备选项、choices、optA；请改成「选项A」「选项B」这类',
      '  —— 名称以「选项」开头的列一律算选项列，包括「选项说明」「选项解析」这种名字；选项列也不会再被当成题干 / 答案 / 题型 / 章节列',
      '• 章节列（可选）：章节 / chapter / 章 / 节 / section / unit；用 > 表示层级，例如「第一章 > 第二节」',
      '• 标签列（可选）：标签 / tags / tag；多个标签用 , ， ; ； | 分隔',
      '• 解析列（可选）：解析 / 答案解析 / 题目解析 / 试题解析 / 分析 / 详解 → 进入题目的"解析"字段，与编辑弹窗里的解析是同一个字段',
      '• 其余列：全部作为附加信息显示在题目下方，并会写进"复制题目"生成的 Callout',
      '⚠ 列名里带「解析 / 说明 / 备注」的列不会被当成题干列或答案列。答案列请直接叫「答案」，不要写成「答案与解析」，否则会提示"未找到答案列"；题干列同理别叫「题目解析」',
      '⚠ 请给每一列都写表头：CSV 中没有表头的列会以空标题出现在附加信息里；Excel 中没有表头的列会被直接丢掉'
    ].forEach(line => fmtLine(fmtBox, line));

    fmtHead(fmtBox, '② 选项怎么写');
    [
      '• 合并成一列、用 | 或换行分隔：A-甲|B-乙|C-丙|D-丁（字母后的分隔符还认 . 、 ． ) ） : ：）',
      '• 合并成一列、用字母标记 + 空格：A.甲 B.乙 C.丙。要求至少 2 个标记且字母顺序递增，否则整段会被当成一个选项',
      '• 分散成多列：选项A / 选项B / 选项C …，每列一个选项',
      '• 项数不限，最多 26 项（受 A–Z 单字母编号限制）',
      '• 合并写法里 | 是分隔符，选项内容本身不能含 |；要放换行或分隔符，请用英文双引号包住整个单元格',
      '• 只有空单元格才算"没有这个选项"。在选项列里，- 、无 、N/A 、\\ 会被当成真实选项内容；但在题干列 / 答案列里它们表示"空值"（该行被跳过）'
    ].forEach(line => fmtLine(fmtBox, line));

    fmtHead(fmtBox, '③ 答案与题型');
    [
      '• 答案只认字母：B / ABD / A,B,D / ＡＢＤ（全角自动转半角）；括号、空格、逗号会被忽略，重复字母自动去重（AAB 等于 AB）',
      '• 判断题答案兼容：对 √ T TRUE 正确 是 Y YES 1 → A；错 × F FALSE 错误 否 N NO 0 → B',
      '⚠ 答案不能写数字序号或 ①②③。写成 1（且没有题型列）会被判成单选题、却没有可判定的正确答案，该题永远算错——请先把序号改成字母',
      '• 没有题型列时自动推断：答案含 2 个以上字母 → 多选题；答案是判断类写法 → 判断题；其余 → 单选题',
      '⚠ 只有 2 个选项、且两个选项文字里分别出现「正确」和「错误」（或「对」「错」）时，会被自动判成判断题。若它其实是 2 选项单选题，请在题型列写明「单选题」',
      '• 判断题没有给选项时，会自动补上 A. 正确 / B. 错误'
    ].forEach(line => fmtLine(fmtBox, line));

    fmtHead(fmtBox, '④ 文件与编码');
    [
      '• 支持 .csv / .tsv / .txt / .text / .xls / .xlsx；可以拖到虚线区，也可以点击虚线区选择文件',
      '• 分隔符（Tab 还是逗号）只看表头那一行，所以表头里不要同时出现 Tab 和逗号',
      '• 字段可以用英文双引号包住，从而包含换行、逗号或 Tab；引号内的 "" 表示一个引号字符',
      '• 默认按 UTF-8 读取；若乱码字符超过 5% 会自动改用 GBK 重读',
      '• Excel 只读取第一个工作表；表头必须在第一行（上面有标题行或空行会错位，请先删掉）',
      '• Excel 中"整列为空"和"没有表头"的列会被丢掉，其余列全部保留',
      '• Excel 日期列按单元格的显示文本读取（不会变成 45123 这类序列号）；显示格式跟随单元格设置，建议统一',
      '• 题干里的 Tab 会替换成空格；题干可以多行（被引号包裹的换行会保留）'
    ].forEach(line => fmtLine(fmtBox, line));

    fmtHead(fmtBox, '⑤ 导入到已有题库 / 数据安全');
    [
      '• 导入到已有题库时，题干重复（忽略空格与大小写）会弹窗让你选择：复制导入（两份并存）/ 替换导入（保留原题 ID，答题记录、错题、收藏的关联不丢）/ 忽略导入（只加新题）/ 取消整个导入',
      '• 题干或答案为空的行会被静默跳过：预览只统计成功解析的题数，不会逐行提示是哪几行',
      '• 题库数据存放在 .obsidian/plugins/quiz-bank/data/ 下（每个题库一个 bank_<id>.json 和一个 bank_<id>_records.json）',
      '• 从 JSON 备份导入前会自动先写一份 backup_<时间戳>.json；换电脑或重装插件前，请用设置里的「导出数据」自己留一份'
    ].forEach(line => fmtLine(fmtBox, line));

    // 预览区域
    const previewEl = contentEl.createEl('div', { cls: 'quiz-import-preview' });
    previewEl.hidden = true;

    // 消息
    const msgEl = contentEl.createEl('div', { cls: 'quiz-import-msg' });

    // ===== 拖拽事件 =====
    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.addClass('dragover');
    });
    dropzone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.removeClass('dragover');
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.removeClass('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        this.readFile(files[0], csvArea, fileInfo, msgEl, nameInput);
      }
    });
    dropzone.addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.csv,.tsv,.txt,.text,.xls,.xlsx';
      input.onchange = (e) => {
        const file = e.target.files[0];
        if (file) {
          this.readFile(file, csvArea, fileInfo, msgEl, nameInput);
        }
      };
      input.click();
    });

    // 按钮区域
    const btnRow = contentEl.createEl('div', { cls: 'quiz-import-btnrow' });

    const previewBtn = btnRow.createEl('button', { cls: 'quiz-btn', text: '🔍 预览解析结果' });
    const importBtn = btnRow.createEl('button', { cls: 'quiz-btn quiz-btn-primary', text: '📥 确认导入' });

    // 预览
    previewBtn.addEventListener('click', () => {
      const content = csvArea.value.trim();
      if (!content) {
        msgEl.className = 'quiz-import-error';
        msgEl.textContent = '请先粘贴内容或拖入文件';
        return;
      }
      this.showPreview(content, previewEl, msgEl);
    });

    // 导入
    importBtn.addEventListener('click', async () => {
      const name = nameInput ? nameInput.value.trim() : '';
      const content = csvArea.value.trim();
      if (!this.targetBank && !name) {
        msgEl.className = 'quiz-import-error';
        msgEl.textContent = '请输入题库名称';
        return;
      }
      if (!content) {
        msgEl.className = 'quiz-import-error';
        msgEl.textContent = '请粘贴内容或拖入文件';
        return;
      }
      try {
        const { questions, errors } = TableParser.parseToQuestions(content);
        if (questions.length === 0) {
          msgEl.className = 'quiz-import-error';
          msgEl.textContent = '未解析到有效题目';
          return;
        }
        importBtn.disabled = true;
        importBtn.textContent = '导入中...';
        if (this.targetBank) {
          // v2.5: 导入题目到已有题库（重复题可选 复制/替换/忽略）
          await this._importToTargetBank(questions, msgEl, importBtn);
          return;
        }
        await this.plugin.createBank(name, questions);
        await this.plugin.saveDataDebounced();
        msgEl.className = 'quiz-import-success';
        msgEl.textContent = `✅ 成功导入 ${questions.length} 题`;
        if (errors.length > 0) msgEl.textContent += `（${errors.length} 行跳过）`;
        new Notice(`题库「${name}」导入成功，共 ${questions.length} 题`);
        this.view.render();
        setTimeout(() => this.close(), 1500);
      } catch (e) {
        msgEl.className = 'quiz-import-error';
        msgEl.textContent = `❌ 导入失败: ${e.message}`;
        importBtn.disabled = false;
        importBtn.textContent = '📥 确认导入';
      }
    });
  }

  /**
   * v2.5: 将解析出的题目导入到已有题库（重复题 = 题干规范化后相同）
   * - 无重复：直接全部追加
   * - 有重复：弹窗选择 复制导入 / 替换导入 / 忽略导入（新题始终导入）
   */
  async _importToTargetBank(questions, msgEl, importBtn) {
    const bank = this.plugin.banks[this.targetBank.id];
    if (!bank) {
      new Notice('题库已不存在，无法导入');
      this.close();
      return;
    }
    const existingMap = new Map();
    for (const q of bank.questions) {
      const key = normalizeQuestionText(q.text);
      if (!existingMap.has(key)) existingMap.set(key, q);
    }
    const fresh = questions.filter(q => !existingMap.has(normalizeQuestionText(q.text)));
    const dupCount = questions.length - fresh.length;

    // 无重复：直接追加
    if (dupCount === 0) {
      bank.questions.push(...questions);
      this.plugin.markBankDirty(bank.id);
      this.plugin.invalidateCache();
      await this.plugin.saveDataDebounced();
      msgEl.className = 'quiz-import-success';
      msgEl.textContent = `✅ 成功导入 ${questions.length} 题（无重复）`;
      new Notice(`成功导入 ${questions.length} 题到「${bank.name}」`);
      this.view.render();
      setTimeout(() => this.close(), 1500);
      return;
    }

    // 有重复：让用户选择处理方式
    importBtn.disabled = false;
    importBtn.textContent = '📥 确认导入';
    new DuplicateImportModal(this.app, dupCount, fresh.length, async (action) => {
      if (!action) { new Notice('已取消导入'); return; }
      let added = 0, copied = 0, replaced = 0, ignored = 0;
      for (const q of questions) {
        const existing = existingMap.get(normalizeQuestionText(q.text));
        if (!existing) {
          bank.questions.push(q);
          existingMap.set(normalizeQuestionText(q.text), q);
          added++;
          continue;
        }
        if (action === 'copy') {
          // 复制导入：重复题作为新题加入（保留两份）
          bank.questions.push(q);
          copied++;
        } else if (action === 'replace') {
          // 替换导入：用导入内容覆盖原题（保留原题 ID，答题记录/错题/收藏关联不丢失）
          existing.text = q.text;
          existing.type = q.type;
          existing.options = q.options;
          existing.answer = q.answer;
          existing.analysis = q.analysis || '';
          existing.tags = q.tags || [];
          existing.extraFields = q.extraFields || {};
          if (q.chapterPath && q.chapterPath.length > 0) existing.chapterPath = q.chapterPath;
          replaced++;
        } else {
          ignored++;
        }
      }
      this.plugin.markBankDirty(bank.id);
      this.plugin.invalidateCache();
      await this.plugin.saveDataDebounced();
      const parts = [`新增 ${added} 题`];
      if (copied > 0) parts.push(`复制导入 ${copied} 题`);
      if (replaced > 0) parts.push(`替换导入 ${replaced} 题`);
      if (ignored > 0) parts.push(`忽略 ${ignored} 题`);
      const summary = parts.join('，');
      msgEl.className = 'quiz-import-success';
      msgEl.textContent = `✅ 导入完成：${summary}`;
      new Notice(`导入完成：${summary}`);
      this.view.render();
      setTimeout(() => this.close(), 2500);
    }).open();
  }

  readFile(file, csvArea, fileInfo, msgEl, nameInput) {
    const validExts = ['.csv', '.tsv', '.txt', '.text', '.xls', '.xlsx'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) {
      msgEl.className = 'quiz-import-error';
      msgEl.textContent = `不支持的文件格式: ${ext}，请使用 .csv/.tsv/.txt/.xls/.xlsx`;
      return;
    }

    // v2.4+: xls/xlsx 文件使用内置解析器转换为 TSV
    if (ext === '.xlsx' || ext === '.xls') {
      this._readXlsxFile(file, csvArea, fileInfo, msgEl, nameInput);
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const buffer = e.target.result;
      let content = new TextDecoder('utf-8', { fatal: false }).decode(buffer);
      const replacementCount = (content.match(/\uFFFD/g) || []).length;
      if (replacementCount > content.length * 0.05) {
        content = new TextDecoder('gbk', { fatal: false }).decode(buffer);
      }
      csvArea.value = content;
      this._fileContent = content;
      this._fileName = file.name;
      // 自动提取文件名（不含路径和扩展名）作为题库名
      if (nameInput) {
        const baseName = file.name.replace(/\.[^/.]+$/, '');
        nameInput.value = baseName;
      }
      fileInfo.hidden = false;
      fileInfo.empty();
      fileInfo.createEl('span', { text: `📄 ${file.name}` });
      fileInfo.createEl('span', { text: `(${(file.size / 1024).toFixed(1)} KB)` });
      const removeBtn = fileInfo.createEl('span', { cls: 'quiz-import-file-remove', text: '✕' });
      removeBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        csvArea.value = '';
        this._fileContent = null;
        this._fileName = null;
        if (nameInput) nameInput.value = '';
        fileInfo.hidden = true;
      });
      msgEl.className = 'quiz-import-success';
      msgEl.textContent = `已加载文件: ${file.name}`;
    };
    reader.onerror = () => {
      msgEl.className = 'quiz-import-error';
      msgEl.textContent = '文件读取失败';
    };
    reader.readAsArrayBuffer(file);
  }

  // v2.7.2: 读取 xls/xlsx 文件，过滤无关列后转为 TSV（async：可能需兜底加载解析库）
  async _readXlsxFile(file, csvArea, fileInfo, msgEl, nameInput) {
    // 库缺失时先尝试兜底加载（从插件目录读取库文件），而不是直接报错
    if (!XLSX || typeof XLSX.read !== 'function') {
      msgEl.className = 'quiz-import-msg';
      msgEl.textContent = '正在加载 Excel 解析库…';
      try {
        await ensureXLSXLoaded(this.plugin, this.plugin.pluginDir || (this.plugin.manifest && this.plugin.manifest.dir));
      } catch (e) { /* 下方统一处理 */ }
    }

    // 仍不可用：给出明确指引，而不是抛 TypeError
    if (!XLSX || typeof XLSX.read !== 'function') {
      msgEl.className = 'quiz-import-error';
      msgEl.textContent = 'Excel 解析库未加载，无法读取 Excel 文件。请改用 CSV（UTF-8 逗号分隔）导入。';
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // 将 sheet 转为二维数组以便过滤列
        // v2.7.4: raw:false —— 取单元格的"显示文本"而不是原始值。
        // 否则日期列会变成 45123 这类序列号（题库里常常带日期列），带格式的数字也会走样。
        const sheetArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '', blankrows: false, raw: false });
        if (!sheetArray || sheetArray.length < 2) {
          throw new Error('Sheet 为空或无数据行');
        }

        const headers = sheetArray[0].map(h => {
          const s = String(h || '').trim();
          return s === 'None' ? '' : s;
        });

        // v2.7.3: 不再按白名单裁列。
        // TableParser 自身会把「题干/答案/题型/选项」识别为核心列直接消费，
        // 把「章节」识别为章节树，其余列一律逐单元格放进解析（extraFields）。
        // 因此这里只需剔除「无表头」和「整列为空」的列，其余全量透传，
        // 避免像旧白名单那样把用户的附加列（如「题目依据」「内部分类」）丢掉。
        const keepIndices = [];
        for (let i = 0; i < headers.length; i++) {
          if (!headers[i]) continue;                  // 无表头：无法命名，跳过
          const colHasValue = sheetArray.slice(1).some(row => {
            const v = row[i];
            return v !== null && v !== undefined && String(v).trim() !== '';
          });
          if (!colHasValue) continue;                 // 整列为空：跳过
          keepIndices.push(i);
        }

        // 用过滤后的列重建工作表和 TSV
        const filteredData = sheetArray.map(row =>
          keepIndices.map(i => {
            const v = row[i];
            return (v === null || v === undefined) ? '' : v;
          })
        );
        const newSheet = XLSX.utils.aoa_to_sheet(filteredData);
        const tsv = XLSX.utils.sheet_to_csv(newSheet, { FS: '\t', RS: '\n' });

        csvArea.value = tsv;
        this._fileContent = tsv;
        this._fileName = file.name;
        if (nameInput) {
          const baseName = file.name.replace(/\.[^/.]+$/, '');
          nameInput.value = baseName;
        }
        fileInfo.hidden = false;
        fileInfo.empty();
        fileInfo.createEl('span', { text: `📊 ${file.name}` });
        fileInfo.createEl('span', { text: `(${(file.size / 1024).toFixed(1)} KB)` });
        const removeBtn = fileInfo.createEl('span', { cls: 'quiz-import-file-remove', text: '✕' });
        removeBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          csvArea.value = '';
          this._fileContent = null;
          this._fileName = null;
          if (nameInput) nameInput.value = '';
          fileInfo.hidden = true;
        });
        msgEl.className = 'quiz-import-success';
        msgEl.textContent = `已加载 Excel 文件: ${file.name}（已过滤无关列并转为 TSV）`;
      } catch (err) {
        msgEl.className = 'quiz-import-error';
        msgEl.textContent = `Excel 解析失败: ${err.message}。请将 Excel 另存为 CSV（Tab 分隔）后重试。`;
      }
    };
    reader.onerror = () => {
      msgEl.className = 'quiz-import-error';
      msgEl.textContent = '文件读取失败';
    };
    reader.readAsArrayBuffer(file);
  }

  /**
   * v2.7.4: 把随插件分发的 Excel 模板保存到本地。
   * 模板位置：<插件目录>/题库导入模板.xlsx（兼容曾经放在 <插件目录>/assets/ 的版本）。
   * 与设置里「导出数据」用同一套写法：adapter 读二进制 → Blob → a[download]。
   */
  async downloadTemplate() {
    const adapter = this.app && this.app.vault && this.app.vault.adapter;
    const dir = this.plugin.pluginDir || (this.plugin.manifest && this.plugin.manifest.dir) || '';
    const candidates = [
      `${dir}/题库导入模板.xlsx`,
      `${dir}/assets/题库导入模板.xlsx`
    ].filter(p => !!p);
    let buffer = null;
    let used = '';
    try {
      if (adapter && typeof adapter.readBinary === 'function') {
        for (const path of candidates) {
          try {
            const exists = typeof adapter.exists === 'function' ? await adapter.exists(path) : true;
            if (!exists) continue;
            buffer = await adapter.readBinary(path);
            used = path;
            break;
          } catch (e) { /* 换下一个候选路径 */ }
        }
      }
    } catch (e) {
      console.error('[QuizBank] 读取模板失败：', e);
    }
    if (!buffer) {
      new Notice('未找到模板文件（预期位置：插件目录/题库导入模板.xlsx），请重新安装插件', 6000);
      return;
    }
    try {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '题库导入模板.xlsx';
      if (document.body && typeof document.body.appendChild === 'function') document.body.appendChild(a);
      a.click();
      if (typeof a.remove === 'function') a.remove();
      setTimeout(() => { try { URL.revokeObjectURL(url); } catch (e) { /* noop */ } }, 3000);
      new Notice('模板已开始下载：题库导入模板.xlsx');
    } catch (e) {
      console.error('[QuizBank] 下载模板失败：', e);
      new Notice('下载失败，可手动复制插件目录下的：' + used, 6000);
    }
  }

  showPreview(content, previewEl, msgEl) {
    try {
      const { questions, errors, delimiter } = TableParser.parseToQuestions(content);
      previewEl.hidden = false;
      previewEl.empty();
      const sepName = delimiter === '\t' ? 'Tab(TSV)' : '逗号(CSV)';
      const summary = previewEl.createEl('div', { cls: 'quiz-import-success' });
      summary.textContent = `✅ 成功解析 ${questions.length} 题 | 分隔符: ${sepName}`;
      if (errors.length > 0) summary.textContent += ` | ${errors.length} 行有误`;
      questions.slice(0, 5).forEach((q, i) => {
        const row = previewEl.createEl('div', { cls: 'quiz-import-preview-row' });
        row.createEl('strong', { text: `${i + 1}. ` });
        row.createSpan({ text: `[${TYPE_LABELS[q.type]}] ${q.text.substring(0, 60)}${q.text.length > 60 ? '...' : ''}` });
        row.createEl('br');
        const meta = row.createEl('span', { text: `   答案: ${q.answer} | 选项: ${q.options.length}个` });
        // v2.4: 显示章节信息
        if (q.chapterPath && q.chapterPath.length > 0) {
          meta.textContent += ` | 章节: ${q.chapterPath.join(' > ')}`;
        }
        if (Object.keys(q.extraFields).length > 0) {
          meta.textContent += ` | 附加: ${Object.keys(q.extraFields).join(', ')}`;
        }
      });
      if (questions.length > 5) {
        previewEl.createEl('div', { cls: 'quiz-import-preview-row', text: `... 还有 ${questions.length - 5} 题` });
      }
      if (errors.length > 0) {
        const errDiv = previewEl.createEl('div', { cls: 'quiz-import-error quiz-import-errors' });
        errors.slice(0, 5).forEach(err => errDiv.createEl('div', { text: `⚠️ ${err}` }));
        if (errors.length > 5) errDiv.createEl('div', { text: `... 还有 ${errors.length - 5} 个错误` });
      }
      msgEl.textContent = '';
    } catch (e) {
      previewEl.hidden = true;
      msgEl.className = 'quiz-import-error';
      msgEl.textContent = `❌ 解析失败: ${e.message}`;
    }
  }
  onClose() { this.contentEl.empty(); }
}

/* ============================================================
   Section 12: 组卷模态框（v2.6：答案类型筛选 + 按类型精确抽题）
              v2.7.5：判断题支持 ✕（错误）/ ✓（正确）双向筛选与精确抽题
   ============================================================ */

class GroupPaperModal extends Modal {
  constructor(app, plugin, view) {
    super(app);
    this.plugin = plugin;
    this.view = view;
    // v2.5: 筛选状态（切换章节后保留）
    this.chapterFilter = 'all';               // 'all' | 具体章节路径
    this.chapterFilterQuestions = null;       // 当前章节筛选下的题目集合
    this.selectedMultiCategories = new Set(); // 勾选的多选答案类型（空 = 不限制）
    // v2.7.5: 判断题答案筛选改为可选集合 —— 'wrong' = ✕（错误），'right' = ✓（正确）
    // 默认只勾 'wrong'，与旧版「只抽错误」的行为一致；两个都勾 = 不限制
    this.judgeAnswerFilters = new Set(['wrong']);
    this.multiCatDrawInputs = {};             // v2.6: 各答案类型的精确抽取数量输入框
    this.judgeDrawInputs = {};                // v2.7.5: 判断题 ✕ / ✓ 各自的精确抽取数量输入框
  }

  /** 多选题答案类型分类：「全选」/「少选 N 项」 */
  getMultiCategory(q) {
    const total = (q.options || []).length;
    const correct = (q.answer || '').length;
    const diff = total - correct;
    return diff <= 0 ? '全选' : `少选 ${diff} 项`;
  }

  /**
   * v2.7.5: 统一判定判断题答案方向。
   * 不能直接比较 q.answer 与选项 key —— 题库里答案可能还停留在「正确 / 错误」这类文字上
   * （normalizeQuestionData 的严格正则只认单字「错」，两个字的「错误」不会被转成 B），
   * 所以改走判分统一入口 getCorrectAnswers()，把答案先归一成选项 key 再比对。
   * @returns {'wrong'|'right'|null} null 表示答案无法判定
   */
  judgeAnswerKey(q) {
    const { wrongKey, rightKey } = judgeOptionKeys(q);
    if (wrongKey === rightKey) return null;   // 数据异常，不强行归类
    const correct = getCorrectAnswers(q);
    if (!correct.length) return null;
    if (correct.includes(wrongKey)) return 'wrong';
    if (correct.includes(rightKey)) return 'right';
    return null;
  }

  /** 判断题答案是否为「✕（错误）」 */
  isJudgeWrong(q) {
    return this.judgeAnswerKey(q) === 'wrong';
  }

  /** v2.7.5: 判断题答案是否为「✓（正确）」 */
  isJudgeRight(q) {
    return this.judgeAnswerKey(q) === 'right';
  }

  /** v2.6: 读取多选题各答案类型的精确抽取数量（>0 才生效） */
  getMultiCatDrawCounts() {
    const counts = {};
    for (const [cat, input] of Object.entries(this.multiCatDrawInputs || {})) {
      const n = parseInt(input && input.value) || 0;
      if (n > 0) counts[cat] = n;
    }
    return counts;
  }

  /**
   * v2.7.5: 判断题指定答案（'wrong' = ✕ 错误 / 'right' = ✓ 正确）的精确抽取数量。
   * @param {'wrong'|'right'} key
   * @returns {number} 0 表示不指定
   */
  getJudgeDrawCount(key) {
    const input = this.judgeDrawInputs && this.judgeDrawInputs[key];
    const n = parseInt(input && input.value) || 0;
    return n > 0 ? n : 0;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('quiz-modal');
    contentEl.createEl('h3', { cls: 'quiz-modal-title', text: '📝 智能组卷' });
    const bank = this.view.currentBank;
    if (!bank || bank.questions.length === 0) {
      contentEl.createEl('div', { cls: 'quiz-empty-text', text: '当前题库为空' });
      return;
    }
    this.bank = bank;
    this.hasChapters = bank.questions.some(q => q.chapterPath && q.chapterPath.length > 0);
    this.bodyEl = contentEl.createEl('div');
    this.renderBody();
  }

  /** 计算当前章节筛选下的题目集合 */
  computeChapterQuestions() {
    if (this.chapterFilter === 'all') return this.bank.questions;
    const tree = ChapterTreeBuilder.build(this.bank.questions);
    const findNode = (node) => {
      if (node.path === this.chapterFilter) return node;
      for (const child of Object.values(node.children)) {
        const found = findNode(child);
        if (found) return found;
      }
      return null;
    };
    const target = findNode(tree);
    if (!target) return this.bank.questions;
    const ids = ChapterTreeBuilder.collectQuestionIds(target);
    return this.bank.questions.filter(q => ids.includes(q.id));
  }

  /** 某题型在所有筛选条件叠加后的可用题池 */
  getTypePool(type) {
    let pool = this.chapterFilterQuestions.filter(q => q.type === type);
    if (type === QUESTION_TYPES.MULTIPLE && this.selectedMultiCategories.size > 0) {
      pool = pool.filter(q => this.selectedMultiCategories.has(this.getMultiCategory(q)));
    }
    // v2.7.5: 判断题按勾选的答案（✕ / ✓）限定随机池；两个都勾等于不限制
    if (type === QUESTION_TYPES.JUDGE && this.judgeAnswerFilters.size > 0 && this.judgeAnswerFilters.size < 2) {
      pool = pool.filter(q => this.judgeAnswerFilters.has(this.judgeAnswerKey(q)));
    }
    return pool;
  }

  hasActiveFilter(type) {
    if (type === QUESTION_TYPES.MULTIPLE) return this.selectedMultiCategories.size > 0;
    // 只有一个答案被勾选才算「有筛选」，两个都勾等于不限制
    if (type === QUESTION_TYPES.JUDGE) return this.judgeAnswerFilters.size > 0 && this.judgeAnswerFilters.size < 2;
    return false;
  }

  /** 刷新题型行的「可用 N 题」标签与数量上限（勾选筛选后实时联动） */
  updateTypeRow(type) {
    const entry = this.inputs[type];
    if (!entry) return;
    const pool = this.getTypePool(type);
    const total = entry.baseCount;
    if (this.hasActiveFilter(type)) {
      entry.labelEl.textContent = `${TYPE_LABELS[type]}（可用 ${pool.length} / 共 ${total} 题）`;
    } else {
      entry.labelEl.textContent = `${TYPE_LABELS[type]}（共 ${total} 题）`;
    }
    entry.input.max = String(pool.length);
    const val = parseInt(entry.input.value) || 0;
    if (val > pool.length) entry.input.value = String(pool.length);
  }

  renderBody() {
    if (!this.chapterFilterQuestions) this.chapterFilterQuestions = this.bank.questions;
    this.multiCatDrawInputs = {};
    this.judgeDrawInputs = {};
    const container = this.bodyEl;
    container.empty();
    const bank = this.bank;

    // 章节筛选下拉
    if (this.hasChapters) {
      const chapterRow = container.createEl('div', { cls: 'quiz-group-row' });
      chapterRow.createEl('span', { cls: 'quiz-group-label', text: '章节筛选' });
      const chSelect = chapterRow.createEl('select', { cls: 'quiz-search-select' });
      chSelect.createEl('option', { value: 'all', text: `全部章节（${bank.questions.length} 题）` });
      const tree = ChapterTreeBuilder.build(bank.questions);
      const addOptions = (node, depth) => {
        const children = Object.values(node.children).sort();
        for (const child of children) {
          const prefix = '  '.repeat(depth) + (depth > 0 ? '├ ' : '');
          const opt = chSelect.createEl('option', { value: child.path, text: `${prefix}${child.name}（${child.count} 题）` });
          if (Object.keys(child.children).length > 0) addOptions(child, depth + 1);
        }
      };
      addOptions(tree, 0);
      chSelect.value = this.chapterFilter;
      chSelect.addEventListener('change', () => {
        this.chapterFilter = chSelect.value;
        this.chapterFilterQuestions = this.computeChapterQuestions();
        this.renderBody();
      });
    }

    const typeCounts = {};
    this.chapterFilterQuestions.forEach(q => { typeCounts[q.type] = (typeCounts[q.type] || 0) + 1; });
    this.inputs = {};
    Object.entries(TYPE_LABELS).forEach(([type, label]) => {
      const count = typeCounts[type] || 0;
      if (count === 0) return;

      // v2.5: 多选题——先统计答案类型分布（全选/少选 N 项），并清理章节切换后已不存在的分类
      let multiCats = null, multiCatMap = null;
      if (type === QUESTION_TYPES.MULTIPLE) {
        multiCatMap = {};
        this.chapterFilterQuestions
          .filter(q => q.type === QUESTION_TYPES.MULTIPLE)
          .forEach(q => {
            const cat = this.getMultiCategory(q);
            if (!multiCatMap[cat]) multiCatMap[cat] = 0;
            multiCatMap[cat]++;
          });
        multiCats = Object.keys(multiCatMap).sort((a, b) => {
          const na = a === '全选' ? 0 : (parseInt(a.replace(/\D+/g, ''), 10) || 999);
          const nb = b === '全选' ? 0 : (parseInt(b.replace(/\D+/g, ''), 10) || 999);
          return na - nb;
        });
        for (const c of [...this.selectedMultiCategories]) {
          if (!multiCats.includes(c)) this.selectedMultiCategories.delete(c);
        }
      }

      const row = container.createEl('div', { cls: 'quiz-group-row' });
      const labelEl = row.createEl('span', { cls: 'quiz-group-label' });
      const input = row.createEl('input', { cls: 'quiz-group-input', attr: { type: 'number', min: '0', max: String(count), value: '0' } });
      this.inputs[type] = { input, labelEl, baseCount: count };
      this.updateTypeRow(type);

      // v2.6: 多选题答案类型筛选行（勾选 = 限定随机池；填写数量 = 按类型精确抽取）
      if (type === QUESTION_TYPES.MULTIPLE && multiCats) {
        const filterRow = container.createEl('div', { cls: 'quiz-group-filter-row' });
        filterRow.createEl('div', { cls: 'quiz-group-filter-title', text: '答案类型筛选（不勾选 = 不限制；填写抽取数量 = 按该类型精确抽题）' });
        multiCats.forEach(cat => {
          const item = filterRow.createEl('span', { cls: 'quiz-group-cat-item' });
          const chkLabel = item.createEl('label', { cls: 'quiz-group-check' });
          const chk = chkLabel.createEl('input', { attr: { type: 'checkbox' } });
          chk.checked = this.selectedMultiCategories.has(cat);
          chkLabel.createEl('span', { text: `答案${cat}` });
          chk.addEventListener('change', () => {
            if (chk.checked) this.selectedMultiCategories.add(cat);
            else this.selectedMultiCategories.delete(cat);
            this.updateTypeRow(QUESTION_TYPES.MULTIPLE);
          });
          const drawInput = item.createEl('input', {
            cls: 'quiz-group-draw-input',
            attr: { type: 'number', min: '0', max: String(multiCatMap[cat]), placeholder: '0' }
          });
          drawInput.title = `从「答案${cat}」中抽取的数量（共 ${multiCatMap[cat]} 题，0 或留空 = 不指定）`;
          item.createEl('span', { cls: 'quiz-group-draw-total', text: `题（共 ${multiCatMap[cat]} 题）` });
          this.multiCatDrawInputs[cat] = drawInput;
        });
      }

      // v2.7.5: 判断题答案筛选行 —— ✕（错误）/ ✓（正确）各自可勾选，也各自可填精确抽取数量
      if (type === QUESTION_TYPES.JUDGE) {
        const judgeQs = this.chapterFilterQuestions.filter(q => q.type === QUESTION_TYPES.JUDGE);
        const wrongCount = judgeQs.filter(q => this.isJudgeWrong(q)).length;
        const rightCount = judgeQs.filter(q => this.isJudgeRight(q)).length;
        const filterRow = container.createEl('div', { cls: 'quiz-group-filter-row' });
        filterRow.createEl('div', { cls: 'quiz-group-filter-title', text: '答案筛选（不勾选 = 不限制；填写抽取数量 = 按该答案精确抽题）' });
        const entries = [
          { key: 'wrong', text: '只抽答案为 ✕（错误）的题', total: wrongCount },
          { key: 'right', text: '只抽答案为 ✓（正确）的题', total: rightCount }
        ];
        for (const { key, text, total } of entries) {
          const item = filterRow.createEl('span', { cls: 'quiz-group-cat-item' });
          const chkLabel = item.createEl('label', { cls: 'quiz-group-check' });
          const chk = chkLabel.createEl('input', { attr: { type: 'checkbox' } });
          chk.checked = this.judgeAnswerFilters.has(key);
          chkLabel.createEl('span', { text });
          chk.addEventListener('change', () => {
            if (chk.checked) this.judgeAnswerFilters.add(key);
            else this.judgeAnswerFilters.delete(key);
            this.updateTypeRow(QUESTION_TYPES.JUDGE);
          });
          const drawInput = item.createEl('input', {
            cls: 'quiz-group-draw-input',
            attr: { type: 'number', min: '0', max: String(total), placeholder: '0' }
          });
          drawInput.title = `从「${key === 'wrong' ? '✕（错误）' : '✓（正确）'}」判断题中抽取的数量（共 ${total} 题，0 或留空 = 不指定）`;
          item.createEl('span', { cls: 'quiz-group-draw-total', text: `题（共 ${total} 题）` });
          this.judgeDrawInputs[key] = drawInput;
        }
      }
    });

    const shuffleRow = container.createEl('div', { cls: 'quiz-group-row' });
    shuffleRow.createEl('span', { cls: 'quiz-group-label', text: '打乱题目顺序' });
    const shuffleCheck = shuffleRow.createEl('input', { attr: { type: 'checkbox' } });
    shuffleCheck.checked = true;
    const genBtn = container.createEl('button', { cls: 'quiz-btn quiz-btn-primary quiz-group-gen-btn', text: '🎲 生成试卷' });
    genBtn.addEventListener('click', () => {
      const selected = [];
      const usedIds = new Set();
      // 按池随机抽 n 题（自动跳过本轮已抽过的题，避免与精确抽取重复）
      const drawFrom = (pool, n) => {
        const avail = pool.filter(q => !usedIds.has(q.id));
        const shuffled = [...avail].sort(() => Math.random() - 0.5);
        const picked = shuffled.slice(0, Math.min(n, avail.length));
        picked.forEach(q => usedIds.add(q.id));
        return picked;
      };
      // 1) 多选题：按答案类型精确抽取（全选/少选 N 项各按各自数量）
      const multiCatCounts = this.getMultiCatDrawCounts();
      for (const [cat, n] of Object.entries(multiCatCounts)) {
        const pool = this.chapterFilterQuestions.filter(q => q.type === QUESTION_TYPES.MULTIPLE && this.getMultiCategory(q) === cat);
        selected.push(...drawFrom(pool, n));
      }
      // 2) 判断题：✕（错误）/ ✓（正确）各自按填写的数量精确抽取
      const judgeDraws = [
        { key: 'wrong', test: (q) => this.isJudgeWrong(q) },
        { key: 'right', test: (q) => this.isJudgeRight(q) }
      ];
      for (const { key, test } of judgeDraws) {
        const n = this.getJudgeDrawCount(key);
        if (n <= 0) continue;
        const pool = this.chapterFilterQuestions.filter(q => q.type === QUESTION_TYPES.JUDGE && test(q));
        selected.push(...drawFrom(pool, n));
      }
      // 3) 各题型行的随机抽取（受勾选筛选限制，并自动排除上面已抽的题）
      Object.entries(this.inputs).forEach(([type, entry]) => {
        const count = parseInt(entry.input.value) || 0;
        if (count > 0) selected.push(...drawFrom(this.getTypePool(type), count));
      });
      if (selected.length === 0) {
        new Notice('请至少填写一个抽取数量');
        return;
      }
      if (shuffleCheck.checked) selected.sort(() => Math.random() - 0.5);
      this.close();
      this.view.startPracticeMode(selected, '组卷练习');
      new Notice(`已生成 ${selected.length} 题试卷`);
    });
  }
  onClose() { this.contentEl.empty(); }
}

/* ============================================================
   Section 13: 主视图 QuizView
   ============================================================ */

class QuizView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf);
    this._initFields(plugin);
  }

  /** 初始化全部实例字段（单一来源：桌面 ItemView 与安卓 Modal 共用，避免新增字段漏初始化） */
  _initFields(plugin) {
    this.plugin = plugin;
    this.currentBank = null;
    this.currentQuestions = [];
    this.currentIndex = 0;
    this.selectedOptions = new Set();
    this.hasAnswered = false;
    this.answerRecords = {};
    this.mode = 'list';
    this.modeLabel = '';
    this.isModalHost = false;      // 安卓端由 QuizMobileModal 置为 true（无真实 leaf）
    this._swipeBoundEl = null;     // 已绑定滑动手势的容器（避免 render 反复注册）
    this.touchStartX = 0;
    this.touchStartY = 0;
    this._keyHandler = null;
    this._contextMenu = null;
    this._closeMenuHandler = null;
    this._closeMenuEscHandler = null;
    // v2.4: 章节相关
    this.chapterTree = null;           // 当前题库的章节树
    this.expandedChapters = new Set(); // 展开的章节路径
    this.selectedChapterIds = new Set(); // 选中的章节叶节点问题ID（空=全选）
    this._chapterFilter = null;        // 章节筛选回调
    // v2.5: 题库列表是否正在查看已隐藏题库
    this.showHiddenBanks = false;
    // v2.6.1: 当前练习是否来自章节选择页（返回时回章节页）
    this.practiceSourceChapter = false;
    // v2.7.6: 布局状态类（答题页 / 安卓平板）的旋转重算监听
    this._layoutResizeBound = false;
    this._layoutResizeHandler = null;
  }

  getViewType() { return VIEW_TYPE_QUIZ; }
  getDisplayText() { return '题库刷题'; }
  getIcon() { return 'book-open'; }

  async onOpen() {
    this.render();
    this.setupKeyboard();
  }

  async onClose() {
    this.cleanupKeyboard();
    this.removeContextMenu();
    this.contentEl.empty();
  }

  getPracticeKey() {
    switch (this.modeLabel) {
      case '刷题模式': return 'practice';
      case '错题重练': return 'wrong';
      case '收藏练习': return 'favorites';
      case '组卷练习': return 'group';
      case '搜索结果': return 'search';
      default: return 'practice';
    }
  }

  /** v2.7.4: 焦点是否在输入控件/可编辑区域（Obsidian 编辑器是 contenteditable，必须单独判断） */
  isEditableTarget(target) {
    if (!target || typeof target !== 'object') return false;
    if (target.isContentEditable) return true;
    const tag = target.tagName;
    return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
  }

  /** v2.7.4: 当前视图是否为键盘操作的合法目标（安卓端全屏 Modal 无真实 leaf，直接放行） */
  isKeyboardTargetActive() {
    if (this.isModalHost) return true;
    try {
      const ws = this.app && this.app.workspace;
      if (!ws) return true;
      if (ws.activeLeaf && this.leaf) return ws.activeLeaf === this.leaf;
      if (typeof ws.getActiveViewOfType === 'function') {
        const active = ws.getActiveViewOfType(QuizView);
        return active ? active === this : false;
      }
    } catch (e) { /* 判断失败时不阻断快捷键 */ }
    return true;
  }

  setupKeyboard() {
    this._keyHandler = (e) => {
      if (this.mode === 'list' || this.mode === 'search' || this.mode === 'chapter-select') return;
      // v2.7.4: 原来不判断焦点，导致刷题视图在后台标签页/分屏里也吃键盘 ——
      // 在笔记里打字时按 s 会弹答题卡、按 A-Z 会静默选中并提交选项。
      if (!this.isKeyboardTargetActive()) return;
      if (this.isEditableTarget(e.target)) return;
      const questions = this.currentQuestions;
      if (!questions || questions.length === 0) return;
      switch (e.key) {
        case 'ArrowLeft': e.preventDefault(); this.prevQuestion(); break;
        case 'ArrowRight': e.preventDefault(); this.nextQuestion(); break;
        case 'Enter':
          e.preventDefault();
          if (this.hasAnswered) { this.nextQuestion(); }
          else {
            const q = questions[this.currentIndex];
            if (q && q.type === QUESTION_TYPES.MULTIPLE) this.submitAnswer();
          }
          break;
        case 'r': case 'R': if (this.hasAnswered) this.retryQuestion(); break;
        case 's': case 'S': e.preventDefault(); new AnswerSheetModal(this.app, this).open(); break;
        default:
          // 支持 A-Z 全部字母选项（选项数超过 6 个时也可键盘作答）
          if (/^[a-z]$/i.test(e.key)) this.handleKeyOption(e.key.toUpperCase());
          break;
      }
    };
    document.addEventListener('keydown', this._keyHandler);
  }

  cleanupKeyboard() {
    if (this._keyHandler) {
      document.removeEventListener('keydown', this._keyHandler);
      this._keyHandler = null;
    }
  }

  handleKeyOption(key) {
    const q = this.currentQuestions[this.currentIndex];
    if (!q || this.hasAnswered) return;
    const validKeys = q.options.map(o => o.key);
    if (!validKeys.includes(key)) return;
    if (q.type === QUESTION_TYPES.SINGLE || q.type === QUESTION_TYPES.JUDGE) {
      this.selectedOptions.clear();
      this.selectedOptions.add(key);
      this.submitAnswer();
    } else {
      if (this.selectedOptions.has(key)) this.selectedOptions.delete(key);
      else this.selectedOptions.add(key);
      this.refreshOptionDisplay();
    }
  }

  setupTouchSwipe(container) {
    // v2.7.4: render() 每次都会调用本方法，而 container 始终是同一个 contentEl。
    // 原先每次 render 都 addEventListener，答几十题后累积几十个监听器：
    // 一次滑动会触发 N 次 nextQuestion（直接跳过 N 题），同时是内存泄漏。
    if (!container || this._swipeBoundEl === container) return;
    this._swipeBoundEl = container;
    container.addEventListener('touchstart', (e) => {
      this.touchStartX = e.touches[0].clientX;
      this.touchStartY = e.touches[0].clientY;
    }, { passive: true });
    container.addEventListener('touchend', (e) => {
      if (this.mode === 'list' || this.mode === 'search' || this.mode === 'chapter-select') return;
      const dx = e.changedTouches[0].clientX - this.touchStartX;
      const dy = e.changedTouches[0].clientY - this.touchStartY;
      if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.5) {
        if (dx < 0) this.nextQuestion();
        else this.prevQuestion();
        hapticFeedback();
      }
    }, { passive: true });
  }

  async render() {
    const container = this.contentEl;
    container.empty();
    container.addClass('quiz-container');
    // 添加平板类标识，用于 CSS 媒体查询之外的 JS 控制（如果需要）
    this._syncLayoutClasses(container);

    // v2.5: 应用答题界面字体缩放
    this.applyFontScale();

    switch (this.mode) {
      case 'list': await this.renderBankList(container); break;
      case 'chapter-select': this.renderChapterSelect(container); break;
      case 'practice': case 'review': case 'wrong': case 'favorites': case 'group':
        this.renderQuestionView(container); break;
      case 'search': await this.renderSearchView(container); break;
    }
    this.setupTouchSwipe(container);
  }

  /**
   * v2.7.6: 同步布局用的状态类。
   *
   * 1) quiz-question-view —— 答题页标识。
   *    原来 CSS 用 :has(> .quiz-question-card) 判断，但 Android WebView 要
   *    Chrome 105+ 才支持；旧机型上整组规则失效，题干会铺满整屏、行长失控。
   *    改由 JS 打类，与 WebView 版本无关。
   * 2) quiz-android-tablet —— 安卓平板（≥600px）标识。
   *    原来只在 render 时判断一次，旋转屏幕后不会更新；现在旋转时也重算。
   */
  _syncLayoutClasses(container) {
    container.classList.toggle('quiz-question-view', isQuestionMode(this.mode));
    container.classList.toggle('quiz-android-tablet', isAndroid() && window.innerWidth >= 600);

    // 旋转 / 改窗口大小后重新判断平板标识（只切类，不重渲染）
    if (this._layoutResizeBound) return;
    this._layoutResizeBound = true;
    const onResize = debounce(() => {
      const el = this.contentEl;
      if (!el || !el.isConnected) return;
      el.classList.toggle('quiz-android-tablet', isAndroid() && window.innerWidth >= 600);
    }, 300);
    if (this.isModalHost) {
      // 全屏 Modal：只需在关闭时移除
      this._layoutResizeHandler = onResize;
      window.addEventListener('resize', onResize, { passive: true });
    } else {
      // ItemView：随视图卸载自动清理
      this.registerDomEvent(window, 'resize', onResize);
    }
  }

  async renderBankList(container) {
    const nav = container.createEl('div', { cls: 'quiz-nav' });
    const navLeft = nav.createEl('div', { cls: 'quiz-nav-left' });
    navLeft.createEl('span', { cls: 'quiz-nav-title', text: '📚 题库刷题' });

    const navRight = nav.createEl('div', { cls: 'quiz-nav-right' });
    const importBtn = navRight.createEl('button', { cls: 'quiz-nav-btn', text: '📥' });
    importBtn.title = '导入题库';
    importBtn.setAttribute('aria-label', '导入题库');
    importBtn.addEventListener('click', () => new ImportModal(this.app, this.plugin, this).open());

    const searchBtn = navRight.createEl('button', { cls: 'quiz-nav-btn', text: '🔍' });
    searchBtn.title = '搜索题目';
    searchBtn.setAttribute('aria-label', '搜索题目');
    searchBtn.addEventListener('click', () => { this.mode = 'search'; this.render(); });

    const banks = await this.plugin.getAllBanks();
    // 防御：跳过结构损坏的题库（无 id / 无 name / 无 questions 数组），避免渲染中断
    const validBanks = banks.filter(b => b && b.id && b.name && Array.isArray(b.questions));

    // v2.5: 隐藏题库——统计数量并提供查看切换
    const hiddenCount = validBanks.filter(b => b.hidden).length;
    // v2.7.4: 没有隐藏题库时不再显示"🙈 0"这个无意义按钮
    if (hiddenCount > 0 || this.showHiddenBanks) {
      const hideToggleBtn = navRight.createEl('button', {
        cls: `quiz-btn quiz-btn-sm${this.showHiddenBanks ? ' quiz-btn-primary' : ''}`,
        text: `🙈 ${hiddenCount}`
      });
      hideToggleBtn.title = this.showHiddenBanks ? '返回题库列表' : `查看已隐藏题库（${hiddenCount} 个）`;
      hideToggleBtn.setAttribute('aria-label', hideToggleBtn.title);
      hideToggleBtn.addEventListener('click', () => { this.showHiddenBanks = !this.showHiddenBanks; this.render(); });
    }

    if (validBanks.length === 0) {
      const empty = container.createEl('div', { cls: 'quiz-empty' });
      empty.createEl('div', { cls: 'quiz-empty-icon', text: '📖' });
      empty.createEl('div', { cls: 'quiz-empty-text', text: '还没有题库，点击右上角 📥 导入' });
      return;
    }

    // v2.5: 按当前视图过滤（正常视图 = 仅未隐藏；隐藏视图 = 仅已隐藏）
    const visibleBanks = this.showHiddenBanks
      ? validBanks.filter(b => b.hidden)
      : validBanks.filter(b => !b.hidden);
    if (visibleBanks.length === 0) {
      const empty = container.createEl('div', { cls: 'quiz-empty' });
      empty.createEl('div', { cls: 'quiz-empty-icon', text: this.showHiddenBanks ? '👁' : '🙈' });
      empty.createEl('div', {
        cls: 'quiz-empty-text',
        text: this.showHiddenBanks ? '没有已隐藏的题库' : '所有题库均已隐藏，点击右上角 🙈 查看已隐藏题库'
      });
      return;
    }

    const listEl = container.createEl('div', { cls: 'quiz-bank-list' });
    for (const bank of visibleBanks) {
      const card = listEl.createEl('div', { cls: `quiz-bank-card${bank.hidden ? ' quiz-bank-card--hidden' : ''}` });
      const info = card.createEl('div', { cls: 'quiz-bank-info' });

      // 题库名称区域（支持编辑）
      const nameContainer = info.createEl('div', { cls: 'quiz-bank-name-container' });
      const nameDisplay = nameContainer.createEl('span', { cls: 'quiz-bank-name', text: bank.name });

      // 编辑名称按钮
      const editNameBtn = nameContainer.createEl('button', { cls: 'quiz-edit-name-btn', text: '✏️' });
      editNameBtn.title = '修改题库名称';
      editNameBtn.setAttribute('aria-label', '修改题库名称');
      editNameBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.startRenameBank(bank, nameContainer, nameDisplay);
      });

      const wrongCount = this.plugin.getWrongQuestions(bank.id).length;
      const favCount = this.plugin.getFavoriteQuestions(bank.id).length;
      const records = this.plugin.getAnswerRecords(bank.id);
      const practicedCount = Object.keys(records).length;
      // v2.4: 检测是否有章节结构
      const hasChapters = this.bankHasChapters(bank);
      const chapterCount = hasChapters
        ? new Set(bank.questions.filter(q => q.chapterPath).map(q => q.chapterPath[0])).size
        : 0;
      const chapterMeta = hasChapters ? ` | 📂 ${chapterCount} 章` : '';
      info.createEl('div', { cls: 'quiz-bank-meta', text: `📚 题库：${bank.questions.length} 题${chapterMeta}` });
      info.createEl('div', { cls: 'quiz-bank-meta quiz-bank-stats', text: `✍️ 练习：已练 ${practicedCount} | 错题 ${wrongCount} | 收藏 ${favCount}` });

      const actions = card.createEl('div', { cls: 'quiz-bank-actions' });
      const startBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-primary', text: '刷题' });
      startBtn.addEventListener('click', async (e) => { e.stopPropagation(); await this.selectBank(bank); await this._enterBank(bank); });
      const reviewBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: '背题' });
      reviewBtn.addEventListener('click', async (e) => { e.stopPropagation(); await this.selectBank(bank); this._enterReview(bank); });
      if (wrongCount > 0) {
        const wrongBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-danger', text: `错题(${wrongCount})` });
        wrongBtn.addEventListener('click', async (e) => { e.stopPropagation(); await this.selectBank(bank); new WrongQuestionsModal(this.app, this.plugin, this).open(); });
      }
      if (favCount > 0) {
        const favBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: `⭐(${favCount})` });
        favBtn.addEventListener('click', async (e) => { e.stopPropagation(); await this.selectBank(bank); new FavoritesModal(this.app, this.plugin, this).open(); });
      }
      const groupBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: '组卷' });
      groupBtn.addEventListener('click', async (e) => { e.stopPropagation(); await this.selectBank(bank); new GroupPaperModal(this.app, this.plugin, this).open(); });
      // v2.5: 导入题目到此题库（重复题可选 复制/替换/忽略）
      const impBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: '📥' });
      impBtn.title = '导入题目到此题库';
      impBtn.setAttribute('aria-label', '导入题目到此题库');
      impBtn.addEventListener('click', (e) => { e.stopPropagation(); new ImportModal(this.app, this.plugin, this, bank).open(); });
      // v2.5: 隐藏/取消隐藏题库
      const hideBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: bank.hidden ? '👁' : '🙈' });
      hideBtn.title = bank.hidden ? '取消隐藏（恢复显示）' : '隐藏题库（列表中不再显示）';
      hideBtn.setAttribute('aria-label', hideBtn.title);
      hideBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        bank.hidden = !bank.hidden;
        this.plugin.markBankDirty(bank.id);
        await this.plugin.saveDataDebounced();
        this.render();
        new Notice(bank.hidden ? `已隐藏「${bank.name}」` : `已取消隐藏「${bank.name}」`);
      });
      const delBtn = actions.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-danger', text: '🗑' });
      delBtn.title = '删除题库';
      delBtn.setAttribute('aria-label', '删除题库');
      delBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const confirmed = await showConfirm(this.app, '删除题库', `确定要删除「${bank.name}」吗？此操作不可撤销。`);
        if (confirmed) { await this.plugin.deleteBank(bank.id); await this.plugin.saveDataDebounced(); new Notice('题库已删除'); this.render(); }
      });
      card.addEventListener('click', async () => { await this.selectBank(bank); await this._enterBank(bank); });
    }
  }

  // v2.4: 进入题库（有章节→章节选择页，无章节→直接进入对应模式）
  async _enterBank(bank) { return this._enterBankMode(bank, false); }

  // v2.4: 进入背题（复用统一进入逻辑，走背题模式）
  _enterReview(bank) { return this._enterBankMode(bank, true); }

  /** 统一的进入题库逻辑：刷题(false) / 背题(true) */
  async _enterBankMode(bank, isReview) {
    if (this.bankHasChapters(bank)) {
      // 动态构建章节树
      this.chapterTree = ChapterTreeBuilder.build(bank.questions);
      this.expandedChapters = new Set();
      this.selectedChapterIds = new Set();
      this.mode = 'chapter-select';
      this._chapterFilter = isReview ? 'review' : null;
      this.render();
    } else if (isReview) {
      this.startReviewMode(bank.questions);
    } else {
      this.startPracticeWithProgress(bank.questions, '刷题模式');
    }
  }

  /** 判断题库是否含章节结构（统一判断，避免多处重复遍历全部题目） */
  bankHasChapters(bank) {
    return !!bank && bank.questions.some(q => q.chapterPath && q.chapterPath.length > 0);
  }

  // 新增：开始重命名题库
  startRenameBank(bank, container, displayEl) {
    container.empty();
    const input = container.createEl('input', { cls: 'quiz-rename-input', value: bank.name });
    input.focus();
    input.select();

    let cancelled = false;
    const save = async () => {
      if (cancelled) return;   // Esc 取消后 blur 仍会触发，这里拦住
      cancelled = true;        // Enter 保存后 blur 也会再来一次，避免重复保存
      const newName = input.value.trim();
      if (newName && newName !== bank.name) {
        await this.plugin.renameBank(bank.id, newName);
        await this.plugin.saveDataDebounced();
        new Notice('题库名称已更新');
      }
      this.render(); // 重新渲染整个列表以恢复结构
    };

    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); save(); }
      if (e.key === 'Escape') { e.preventDefault(); cancelled = true; this.render(); }
    });

    // 失去焦点时保存
    input.addEventListener('blur', () => {
        // 稍微延迟以避免点击其他元素时立即消失导致无法点击
        setTimeout(save, 100);
    });
  }

  // v2.4: 章节选择页
  renderChapterSelect(container) {
    const bank = this.currentBank;
    const tree = this.chapterTree;
    if (!bank || !tree) { this.mode = 'list'; this.render(); return; }

    // 导航栏
    const nav = container.createEl('div', { cls: 'quiz-nav' });
    const navLeft = nav.createEl('div', { cls: 'quiz-nav-left' });
    const backBtn = navLeft.createEl('button', { cls: 'quiz-nav-btn', text: '←' });
    backBtn.setAttribute('aria-label', '返回题库列表');
    backBtn.addEventListener('click', () => { this.mode = 'list'; this._chapterFilter = null; this.render(); });
    navLeft.createEl('span', { cls: 'quiz-nav-title', text: bank.name });
    navLeft.createEl('span', { cls: 'quiz-nav-progress', text: `${bank.questions.length} 题` });

    const navRight = nav.createEl('div', { cls: 'quiz-nav-right' });
    // "全部刷题"快捷按钮
    const allBtn = navRight.createEl('button', { cls: 'quiz-btn quiz-btn-sm quiz-btn-primary', text: '全部刷题' });
    const isReviewFilter = this._chapterFilter === 'review';
    allBtn.addEventListener('click', () => {
      this.practiceSourceChapter = true;
      if (isReviewFilter) this.startReviewMode(bank.questions);
      else this.startPracticeWithProgress(bank.questions, '刷题模式');
    });

    // 统计栏
    const statsEl = container.createEl('div', { cls: 'quiz-stats' });
    const leafCount = tree.count || bank.questions.length;
    const l1Count = Object.keys(tree.children).length;
    [
      { value: l1Count, label: '章/大类' },
      { value: leafCount, label: '总题数' }
    ].forEach(s => {
      const card = statsEl.createEl('div', { cls: 'quiz-stat-card' });
      card.createEl('div', { cls: 'quiz-stat-value', text: String(s.value) });
      card.createEl('div', { cls: 'quiz-stat-label', text: s.label });
    });

    // 章节树
    const treeEl = container.createEl('div', { cls: 'quiz-chapter-tree' });
    this._renderChapterTree(treeEl, tree, 0, bank);

    // 底部浮动操作按钮
    if (this.selectedChapterIds.size > 0) {
      const floatBar = container.createEl('div', { cls: 'quiz-chapter-floatbar' });
      const selectedCount = this.selectedChapterIds.size;
      const selectLabel = floatBar.createEl('span', { cls: 'quiz-chapter-select-label', text: `已选 ${selectedCount} 题` });
      const floatBtn = floatBar.createEl('button', { cls: 'quiz-btn quiz-btn-primary', text: '开始练习' });
      floatBtn.addEventListener('click', () => {
        this.practiceSourceChapter = true;
        const ids = this.selectedChapterIds;
        const questions = bank.questions.filter(q => ids.has(q.id));
        if (isReviewFilter) this.startReviewMode(questions);
        else this.startPracticeWithProgress(questions, '章节练习');
      });
      const clearBtn = floatBar.createEl('button', { cls: 'quiz-btn', text: '清除选择' });
      clearBtn.addEventListener('click', () => {
        this.selectedChapterIds = new Set();
        this.render();
      });
    }
  }

  /** 递归渲染章节树节点 */
  _renderChapterTree(container, node, depth, bank) {
    const children = Object.values(node.children).sort((a, b) => {
      // 按章节编号排序（如果有的话）
      const numA = parseFloat(a.name.match(/^[\d.]+/)?.[0] || '999');
      const numB = parseFloat(b.name.match(/^[\d.]+/)?.[0] || '999');
      return numA - numB;
    });

    for (const child of children) {
      const isExpanded = this.expandedChapters.has(child.path);
      const hasChildren = Object.keys(child.children).length > 0;

      // 行容器
      const row = container.createEl('div', { cls: 'quiz-chapter-row' });
      row.style.paddingLeft = `${12 + depth * 20}px`;

      // 展开/折叠按钮
      const toggle = row.createEl('span', { cls: 'quiz-chapter-toggle', text: isExpanded ? '▼' : '▶' });
      if (!hasChildren) toggle.addClass('quiz-chapter-toggle--leaf');

      // 复选框
      const checkbox = row.createEl('input', { attr: { type: 'checkbox' } });
      checkbox.className = 'quiz-chapter-check';
      // 判断当前节点是否全选
      const nodeIds = ChapterTreeBuilder.collectQuestionIds(child);
      const allSelected = nodeIds.length > 0 && nodeIds.every(id => this.selectedChapterIds.has(id));
      checkbox.checked = allSelected;
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          for (const id of nodeIds) this.selectedChapterIds.add(id);
        } else {
          for (const id of nodeIds) this.selectedChapterIds.delete(id);
        }
        this.render();
      });

      // 章节名称 + 题数
      const label = row.createEl('span', { cls: 'quiz-chapter-label', text: child.name });
      const badge = row.createEl('span', { cls: 'quiz-chapter-badge', text: String(child.count) });
      badge.title = `${child.count} 题`;

      // 行点击：展开/折叠，或直接刷题
      row.addEventListener('click', (e) => {
        if (e.target.tagName === 'INPUT') return; // 让 checkbox 自己处理
        if (hasChildren) {
          if (this.expandedChapters.has(child.path)) {
            this.expandedChapters.delete(child.path);
          } else {
            this.expandedChapters.add(child.path);
          }
          this.render();
        } else {
          // 叶节点：直接开始练习
          this.practiceSourceChapter = true;
          const questions = bank.questions.filter(q => nodeIds.includes(q.id));
          if (this._chapterFilter === 'review') this.startReviewMode(questions);
          else this.startPracticeWithProgress(questions, '章节练习');
        }
      });

      container.appendChild(row);

      // 递归子节点
      if (hasChildren && isExpanded) {
        this._renderChapterTree(container, child, depth + 1, bank);
      }
    }
  }

  async renderSearchView(container) {
    const nav = container.createEl('div', { cls: 'quiz-nav' });
    const navLeft = nav.createEl('div', { cls: 'quiz-nav-left' });
    const backBtn = navLeft.createEl('button', { cls: 'quiz-nav-btn', text: '←' });
    backBtn.setAttribute('aria-label', '返回题库列表');
    backBtn.addEventListener('click', () => { this.mode = 'list'; this.render(); });
    navLeft.createEl('span', { cls: 'quiz-nav-title', text: '搜索题目' });
    const searchBox = container.createEl('div', { cls: 'quiz-search-box' });
    const searchInput = searchBox.createEl('input', { cls: 'quiz-search-input', attr: { placeholder: '输入关键词搜索...', type: 'text' } });
    const scopeSelect = searchBox.createEl('select', { cls: 'quiz-search-select' });
    scopeSelect.createEl('option', { value: 'all', text: '全部' });
    scopeSelect.createEl('option', { value: 'question', text: '题干' });
    scopeSelect.createEl('option', { value: 'options', text: '选项' });
    // v2.7.6: 加类名以便宽屏下与其他搜索结果对齐（原来是无类名的裸 div，无法参与宽度约束）
    const resultsEl = container.createEl('div', { cls: 'quiz-search-results' });
    let searchLimit = 50;   // v2.7.4: 结果不再硬截断，改按需加载
    const doSearch = async () => {
      const query = searchInput.value.trim().toLowerCase();
      const scope = scopeSelect.value;
      resultsEl.empty();
      if (!query) { resultsEl.createEl('div', { cls: 'quiz-empty-text', text: '输入关键词开始搜索' }); return; }
      const banks = await this.plugin.getAllBanks();
      const index = this.plugin.getSearchIndex();
      let results = [];
      for (const bank of banks) {
        const entries = index[bank.id] || [];
        for (const entry of entries) {
          let match = false;
          if (scope === 'all' || scope === 'question') match = entry.textLower.includes(query);
          if (!match && (scope === 'all' || scope === 'options')) match = entry.optionsLower.some(t => t.includes(query));
          if (match) results.push({ question: entry.q, bank });
        }
      }
      if (results.length === 0) { resultsEl.createEl('div', { cls: 'quiz-empty-text', text: '未找到匹配题目' }); return; }
      const renderList = (list) => {
        resultsEl.empty();
        resultsEl.createEl('div', { cls: 'quiz-bank-meta', text: `找到 ${list.length} 个结果` });
        list.slice(0, searchLimit).forEach(({ question, bank }) => {
          const item = resultsEl.createEl('div', { cls: 'quiz-list-item' });
          const header = item.createEl('div', { cls: 'quiz-list-item-header' });
          header.createEl('span', { cls: 'quiz-question-type', text: `${TYPE_LABELS[question.type]} | ${bank.name}` });
          const textEl = header.createEl('div', { cls: 'quiz-list-item-text' });
          safeSetText(textEl, question.text);
          item.createEl('div', { cls: 'quiz-list-item-answer', text: `答案：${question.answer}` });
          item.addEventListener('click', async () => {
            await this.selectBank(bank);
            const idx = bank.questions.findIndex(q => q.id === question.id);
            if (idx === -1) return;
            // v2.7.4: 走统一练习入口。原先直接赋值 currentQuestions，
            // 导致答题记录用的是全库记录、答题卡统计口径与其它入口不一致。
            this.startPracticeMode(bank.questions, '搜索结果', idx);
          });
        });
        if (list.length > searchLimit) {
          const moreBtn = resultsEl.createEl('button', { cls: 'quiz-btn quiz-btn-sm', text: `显示更多（还有 ${list.length - searchLimit} 条）` });
          moreBtn.addEventListener('click', () => { searchLimit += 50; renderList(list); });
        }
      };
      searchLimit = 50;
      renderList(results);
    };
    searchInput.addEventListener('input', debounce(doSearch, 300));
    scopeSelect.addEventListener('change', doSearch);
    resultsEl.createEl('div', { cls: 'quiz-empty-text', text: '输入关键词开始搜索' });
  }

  renderQuestionView(container) {
    const questions = this.currentQuestions;
    if (!questions || questions.length === 0) { container.createEl('div', { cls: 'quiz-empty-text', text: '暂无题目' }); return; }
    const isReview = this.mode === 'review';
    const q = questions[this.currentIndex];
    if (!q) return;

    // 进度条
    const progressBar = container.createEl('div', { cls: 'quiz-progress-bar' });
    progressBar.createEl('div', { cls: 'quiz-progress-fill' }).style.width = `${((this.currentIndex + 1) / questions.length) * 100}%`;

    // 导航
    const nav = container.createEl('div', { cls: 'quiz-nav' });
    const navLeft = nav.createEl('div', { cls: 'quiz-nav-left' });
    const backBtn = navLeft.createEl('button', { cls: 'quiz-nav-btn', text: '←' });
    backBtn.setAttribute('aria-label', '返回题库列表');
    backBtn.addEventListener('click', () => {
      this.saveCurrentProgress();
      // v2.6.1: 来自章节选择页的练习，返回章节页而不是题库列表
      if (this.practiceSourceChapter && this.currentBank && this.bankHasChapters(this.currentBank)) {
        this.practiceSourceChapter = false;
        this.mode = 'chapter-select';
        this.render();
        return;
      }
      this.mode = 'list';
      this.render();
    });
    navLeft.createEl('span', { cls: 'quiz-nav-title', text: this.modeLabel });
    // v2.4: 显示章节面包屑
    if (q.chapterPath && q.chapterPath.length > 0) {
      navLeft.createEl('span', { cls: 'quiz-nav-progress quiz-chapter-breadcrumb', text: q.chapterPath.join(' > ') });
    }
    navLeft.createEl('span', { cls: 'quiz-nav-progress', text: `${this.currentIndex + 1}/${questions.length}` });

    const navRight = nav.createEl('div', { cls: 'quiz-nav-right' });
    const sheetBtn = navRight.createEl('button', { cls: 'quiz-nav-btn', text: '📋' });
    sheetBtn.title = '答题卡 (S)';
    sheetBtn.setAttribute('aria-label', '答题卡 (S)');
    sheetBtn.addEventListener('click', () => new AnswerSheetModal(this.app, this).open());

    // v2.7.6: 用可变状态维护收藏标记。原来把 isFav 当常量捕获，
    // 点一次收藏后按钮变成 ★，但 isFav 仍是 false，再点依然是「添加」，
    // 于是必须切到别的题再切回来才能取消收藏。
    let isFav = this.currentBank && this.plugin.getFavoriteQuestions(this.currentBank.id).includes(q.id);
    const favBtn = navRight.createEl('button', { cls: `quiz-nav-btn${isFav ? ' active' : ''}`, text: isFav ? '★' : '☆' });
    favBtn.setAttribute('aria-label', isFav ? '取消收藏' : '收藏此题');
    favBtn.addEventListener('click', async () => {
      if (!this.currentBank) return;
      if (isFav) this.plugin.removeFavoriteQuestion(this.currentBank.id, q.id);
      else this.plugin.addFavoriteQuestion(this.currentBank.id, q.id);
      isFav = !isFav;
      await this.plugin.saveDataDebounced();
      favBtn.text = isFav ? '★' : '☆';
      favBtn.toggleClass('active', isFav);
      favBtn.setAttribute('aria-label', isFav ? '取消收藏' : '收藏此题');
      hapticFeedback();
    });

    // 复制题目：与字体调节同排；方法内部自动定位题库原题
    const copyBtn = navRight.createEl('button', { cls: 'quiz-nav-btn', text: '⧉' });
    copyBtn.title = '复制题目为 Callout';
    copyBtn.setAttribute('aria-label', '复制题目为 Callout');
    copyBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      hapticFeedback();
      this.copyQuestionToClipboard(q, this.currentBank ? this.currentBank.name : '');
    });

    // v2.5: 答题界面字体大小调节
    const fontDownBtn = navRight.createEl('button', { cls: 'quiz-nav-btn quiz-font-btn', text: 'A-' });
    fontDownBtn.title = '减小字体';
    fontDownBtn.setAttribute('aria-label', '减小字体');
    fontDownBtn.addEventListener('click', () => this.adjustFont(-1));
    const fontUpBtn = navRight.createEl('button', { cls: 'quiz-nav-btn quiz-font-btn', text: 'A+' });
    fontUpBtn.title = '增大字体';
    fontUpBtn.setAttribute('aria-label', '增大字体');
    fontUpBtn.addEventListener('click', () => this.adjustFont(1));

    const moreBtn = navRight.createEl('button', { cls: 'quiz-nav-btn', text: '⋯' });
    moreBtn.setAttribute('aria-label', '更多操作');
    moreBtn.addEventListener('click', (e) => { e.stopPropagation(); this.showMoreMenu(moreBtn, q); });

    // 题目卡片
    const card = container.createEl('div', { cls: 'quiz-question-card' });
    const header = card.createEl('div', { cls: 'quiz-question-header' });
    header.createEl('span', { cls: 'quiz-question-type', text: TYPE_LABELS[q.type] || q.type });
    if (q.tags && q.tags.length > 0) {
      const tagsEl = header.createEl('span', { cls: 'quiz-bank-meta' });
      safeSetText(tagsEl, q.tags.join(' | '));
    }
    const textEl = card.createEl('div', { cls: 'quiz-question-text' });
    safeSetText(textEl, q.text);

    // 选项
    const optionsEl = card.createEl('div', { cls: 'quiz-options' });
    this.renderOptions(optionsEl, q, isReview);

    // 结果
    if (this.hasAnswered || isReview) {
      this.renderResult(card, q, isReview);
    }

    // 操作栏
    this.renderActions(container);
  }

  renderActions(container) {
    const questions = this.currentQuestions;
    const q = questions[this.currentIndex];
    const isReview = this.mode === 'review';
    const actions = container.createEl('div', { cls: 'quiz-actions' });
    const actionsLeft = actions.createEl('div', { cls: 'quiz-actions-left' });
    const actionsRight = actions.createEl('div', { cls: 'quiz-actions-right' });

    const prevBtn = actionsLeft.createEl('button', { cls: 'quiz-btn', text: '← 上一题' });
    prevBtn.disabled = this.currentIndex === 0;
    prevBtn.addEventListener('click', () => this.prevQuestion());

    const nextBtn = actionsRight.createEl('button', { cls: 'quiz-btn', text: '下一题 →' });
    nextBtn.disabled = this.currentIndex === questions.length - 1;
    nextBtn.addEventListener('click', () => this.nextQuestion());

    if (q.type === QUESTION_TYPES.MULTIPLE && !this.hasAnswered && !isReview) {
      const submitBtn = actionsRight.createEl('button', { cls: 'quiz-btn quiz-btn-primary', text: '提交 (Enter)' });
      submitBtn.addEventListener('click', () => this.submitAnswer());
    }

    if (this.hasAnswered && !isReview) {
      const retryBtn = actionsLeft.createEl('button', { cls: 'quiz-btn', text: '🔄 重做 (R)' });
      retryBtn.addEventListener('click', () => this.retryQuestion());
    }

    const record = this.answerRecords[q.id];
    if (record && !record.correct && !isReview) {
      const simBtn = actionsRight.createEl('button', { cls: 'quiz-btn', text: '🔍 相似题' });
      simBtn.addEventListener('click', () => new SimilarQuestionsModal(this.app, this.plugin, this, q).open());
    }
  }

  /** 局部刷新：仅重建选项区/结果区/操作栏，避免整个视图 DOM 重建（千题级题库下的卡顿来源） */
  refreshAfterAnswer() {
    const container = this.contentEl;
    const card = container.querySelector('.quiz-question-card');
    if (!card) { this.render(); return; }
    const q = this.currentQuestions[this.currentIndex];
    if (!q) { this.render(); return; }
    // 重建选项区
    const oldOptions = card.querySelector('.quiz-options');
    if (oldOptions) {
      const newOptions = card.createEl('div', { cls: 'quiz-options' });
      this.renderOptions(newOptions, q, false);
      oldOptions.remove();
    }
    // 结果区：仅在已作答时显示（重做后不显示）
    card.querySelectorAll('.quiz-result, .quiz-analysis').forEach(el => el.remove());
    if (this.hasAnswered) this.renderResult(card, q, false);
    // 重建操作栏
    const oldActions = container.querySelector('.quiz-actions');
    if (oldActions) {
      const actionsParent = oldActions.parentElement;
      this.renderActions(actionsParent);
      oldActions.remove();
    }
  }

  renderOptions(container, question, isReview) {
    container.empty();
    const correctAnswers = getCorrectAnswers(question);
    question.options.forEach(option => {
      const optionEl = container.createEl('div', { cls: 'quiz-option' });
      // 无障碍：选项可聚焦，Enter/Space 等效点击（stopPropagation 避免与全局快捷键冲突）
      optionEl.tabIndex = 0;
      optionEl.setAttribute('role', 'button');
      optionEl.setAttribute('aria-label', `选项 ${option.key}`);
      const isSelected = this.selectedOptions.has(option.key);
      const isCorrect = correctAnswers.includes(option.key);
      if (this.hasAnswered || isReview) {
        optionEl.addClass('disabled');
        optionEl.setAttribute('aria-disabled', 'true');
        optionEl.tabIndex = -1; // 已作答/背题模式选项不再进入 Tab 顺序
        if (isCorrect) optionEl.addClass('correct');
        if (isSelected && !isCorrect) optionEl.addClass('wrong');
      } else {
        if (isSelected) optionEl.addClass('selected');
      }
      optionEl.createEl('span', { cls: 'quiz-option-key', text: option.key });
      const textEl = optionEl.createEl('span', { cls: 'quiz-option-text' });
      safeSetText(textEl, option.text);
      if (!this.hasAnswered) {
        optionEl.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            e.stopPropagation();
            optionEl.click();
          }
        });
        optionEl.addEventListener('click', () => {
          if (isReview) {
            this.selectedOptions.add(option.key);
            this.hasAnswered = true;
            hapticFeedback();
            this.render();
          } else if (question.type === QUESTION_TYPES.SINGLE || question.type === QUESTION_TYPES.JUDGE) {
            this.selectedOptions.clear();
            this.selectedOptions.add(option.key);
            hapticFeedback();
            this.submitAnswer();
          } else {
            if (this.selectedOptions.has(option.key)) this.selectedOptions.delete(option.key);
            else this.selectedOptions.add(option.key);
            hapticFeedback();
            this.refreshOptionDisplay();
          }
        });
      }
      container.appendChild(optionEl);
    });
  }

  refreshOptionDisplay() {
    const optionsEl = this.contentEl.querySelector('.quiz-options');
    if (!optionsEl) return;
    const q = this.currentQuestions[this.currentIndex];
    if (!q) return;
    const optionEls = optionsEl.querySelectorAll('.quiz-option');
    q.options.forEach((option, i) => {
      const el = optionEls[i];
      if (el) el.classList.toggle('selected', this.selectedOptions.has(option.key));
    });
  }

  renderResult(container, question, isReview) {
    const correctAnswers = getCorrectAnswers(question);
    const userAnswers = [...this.selectedOptions].sort();
    const isCorrect = isReview || (userAnswers.length === correctAnswers.length && userAnswers.every(a => correctAnswers.includes(a)));
    const resultEl = container.createEl('div', { cls: `quiz-result ${isCorrect ? 'correct' : 'wrong'}` });
    const titleEl = resultEl.createEl('div', { cls: 'quiz-result-title' });
    if (isReview) {
      safeSetText(titleEl, `📖 正确答案：${correctAnswers.join('') || question.answer}`);
    } else {
      safeSetText(titleEl, isCorrect ? '✅ 回答正确！' : `❌ 回答错误，正确答案：${correctAnswers.join('') || question.answer}`);
    }
    if (!isReview && !isCorrect) {
      resultEl.createEl('div', { text: `你的答案：${userAnswers.join('') || '未作答'}` });
    }
    if (question.analysis) {
      const analysisEl = container.createEl('div', { cls: 'quiz-analysis' });
      analysisEl.createEl('div', { cls: 'quiz-analysis-title', text: '📝 解析' });
      analysisEl.createEl('div', { cls: 'quiz-extra-field-content', text: question.analysis });
    }
    if (question.extraFields && Object.keys(question.extraFields).length > 0) {
      const extraEl = container.createEl('div', { cls: 'quiz-analysis' });
      Object.entries(question.extraFields).forEach(([key, val]) => {
        if (isEmptyValue(val)) return;
        const field = extraEl.createEl('div', { cls: 'quiz-extra-field' });
        field.createEl('div', { cls: 'quiz-extra-field-title', text: key });
        field.createEl('div', { cls: 'quiz-extra-field-content', text: val });
      });
    }
  }

  /**
   * 复制题目为 Callout Markdown（QuizView 实例方法，供导航按钮/更多菜单调用）
   * 始终基于题库中的原题（而非选项乱序后的练习副本）；内部委托给全局工具函数
   */
  copyQuestionToClipboard(question, bankName) {
    const original = (this.currentBank && question)
      ? (this.currentBank.questions.find(x => x.id === question.id) || question)
      : question;
    const bank = bankName || (this.currentBank ? this.currentBank.name : '');
    return copyQuestionToClipboard(original, bank);
  }

  showMoreMenu(anchorEl, question) {
    this.removeContextMenu();
    // 练习模式下 question 可能是选项乱序后的副本；编辑/相似题一律基于原题，避免洗牌顺序写回题库
    const originalQuestion = this.currentBank
      ? (this.currentBank.questions.find(x => x.id === question.id) || question)
      : question;
    const menu = document.createElement('div');
    menu.className = 'quiz-context-menu';
    menu.setAttribute('role', 'menu');
    menu.setAttribute('aria-label', '更多操作');
    const items = [
      { text: '⏮ 从头开始', action: () => {
        this.currentIndex = 0;
        this.resetQuestionState();
        this.saveCurrentProgress();
        this.render();
        new Notice('已回到第 1 题');
      } },
      { text: '📋 复制题目', action: () => this.copyQuestionToClipboard(originalQuestion, this.currentBank && this.currentBank.name) },
      { text: '✏️ 编辑题目', action: () => new EditQuestionModal(this.app, this.plugin, this, originalQuestion).open() },
      { text: '📋 答题卡', action: () => new AnswerSheetModal(this.app, this).open() },
      { text: '🔍 相似题', action: () => new SimilarQuestionsModal(this.app, this.plugin, this, originalQuestion).open() },
      { text: '🗑 删除此题', action: async () => {
        const confirmed = await showConfirm(this.app, '删除题目', '确定要删除这道题吗？');
        if (confirmed) {
          await this.plugin.deleteQuestion(this.currentBank.id, question.id);
          await this.plugin.saveDataDebounced();
          this.currentQuestions = this.currentQuestions.filter(q => q.id !== question.id);
          if (this.currentIndex >= this.currentQuestions.length) this.currentIndex = Math.max(0, this.currentQuestions.length - 1);
          this.render();
          new Notice('题目已删除');
        }
      }}
    ];
    items.forEach(item => {
      const el = menu.createEl('div');
      el.className = 'quiz-context-menu-item';
      el.setAttribute('role', 'menuitem');
      el.tabIndex = 0;
      el.textContent = item.text;
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); e.stopPropagation(); el.click(); }
      });
      el.addEventListener('click', () => { this.removeContextMenu(); item.action(); });
      menu.appendChild(el);
    });
    const rect = anchorEl.getBoundingClientRect();
    menu.style.top = `${rect.bottom + 4}px`;
    menu.style.right = `${window.innerWidth - rect.right}px`;
    document.body.appendChild(menu);
    this._contextMenu = menu;
    this._closeMenuHandler = (e) => {
      if (menu.isConnected && !menu.contains(e.target)) this.removeContextMenu();
    };
    this._closeMenuEscHandler = (e) => {
      if (e.key === 'Escape') this.removeContextMenu();
    };
    setTimeout(() => document.addEventListener('click', this._closeMenuHandler), 0);
    document.addEventListener('keydown', this._closeMenuEscHandler);
  }

  removeContextMenu() {
    if (this._contextMenu) {
      if (this._contextMenu.isConnected) this._contextMenu.remove();
      this._contextMenu = null;
    }
    if (this._closeMenuHandler) {
      document.removeEventListener('click', this._closeMenuHandler);
      this._closeMenuHandler = null;
    }
    if (this._closeMenuEscHandler) {
      document.removeEventListener('keydown', this._closeMenuEscHandler);
      this._closeMenuEscHandler = null;
    }
  }

  submitAnswer() {
    const q = this.currentQuestions[this.currentIndex];
    if (!q || this.hasAnswered) return;
    if (this.selectedOptions.size === 0) { new Notice('请至少选择一个选项'); return; }
    this.hasAnswered = true;
    const correctAnswers = getCorrectAnswers(q);
    const userAnswers = [...this.selectedOptions].sort();
    const isCorrect = userAnswers.length === correctAnswers.length && userAnswers.every(a => correctAnswers.includes(a));
    this.answerRecords[q.id] = { correct: isCorrect, userAnswer: userAnswers.join(''), timestamp: Date.now() };
    if (this.currentBank) {
      this.plugin.setModeAnswerRecord(this.currentBank.id, this.getPracticeKey(), q.id, this.answerRecords[q.id]);
      this.plugin.setAnswerRecord(this.currentBank.id, q.id, this.answerRecords[q.id]);
      if (!isCorrect) {
        this.plugin.addWrongQuestion(this.currentBank.id, q.id);
      } else if (this.plugin.settings.autoRemoveWrongOnCorrect) {
        // 开关开启时：答对自动移出错题集
        this.plugin.removeWrongQuestion(this.currentBank.id, q.id);
      }
      this.plugin.saveDataDebounced();
    }
    hapticFeedback();
    this.refreshAfterAnswer();
    if (isCorrect && this.plugin.settings.autoNextOnCorrect) {
      setTimeout(() => {
        if (this.hasAnswered && this.currentIndex < this.currentQuestions.length - 1) this.nextQuestion();
      }, 800);
    }
  }

  retryQuestion() {
    this.hasAnswered = false;
    this.selectedOptions.clear();
    const q = this.currentQuestions[this.currentIndex];
    if (q && this.answerRecords[q.id]) {
      delete this.answerRecords[q.id];
      if (this.currentBank) {
        this.plugin.removeAnswerRecord(this.currentBank.id, q.id);
        this.plugin.saveDataDebounced();
      }
    }
    this.refreshAfterAnswer();
  }

  prevQuestion() { if (this.currentIndex > 0) { this.saveCurrentProgress(); this.currentIndex--; this.resetQuestionState(); this.render(); } }
  nextQuestion() { if (this.currentIndex < this.currentQuestions.length - 1) { this.saveCurrentProgress(); this.currentIndex++; this.resetQuestionState(); this.render(); } }
  goToQuestion(index) { if (index >= 0 && index < this.currentQuestions.length) { this.saveCurrentProgress(); this.currentIndex = index; this.resetQuestionState(); this.render(); } }

  resetQuestionState() {
    const q = this.currentQuestions[this.currentIndex];
    if (q && this.answerRecords[q.id]) {
      this.hasAnswered = true;
      this.selectedOptions = new Set(this.answerRecords[q.id].userAnswer.split(''));
    } else {
      this.hasAnswered = false;
      this.selectedOptions.clear();
    }
  }

  async selectBank(bank) {
    this.currentBank = bank;
    this.answerRecords = this.plugin.getAnswerRecords(bank.id) || {};
  }

  /** 选项乱序：按设置打乱每题选项顺序并重建答案字母映射（不修改原始题库对象） */
  applyOptionShuffle(questions) {
    if (!questions || !this.plugin.settings.shuffleOptions) return questions;
    return questions.map(q => {
      // v2.7.6: 判断题不参与乱序。
      // 判断题的选项只有「正确 / 错误」两个，属于固定的语义标签而不是可打乱的候选项；
      // 打乱后界面变成「错误 / 正确」，看着别扭，也不符合判断题的作答习惯。
      if (q.type === QUESTION_TYPES.JUDGE) return q;
      if (!q.options || q.options.length < 2) return q;
      const newQ = Object.assign({}, q, { options: q.options.map(o => Object.assign({}, o)) });
      // Fisher-Yates 洗牌
      for (let i = newQ.options.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const tmp = newQ.options[i];
        newQ.options[i] = newQ.options[j];
        newQ.options[j] = tmp;
      }
      // v2.7.4: 原 key -> 新 key 映射必须取「该选项自己的原 key」。
      // 原来取的是「洗牌前位于同一位置的 key」(origKeys[idx])，而新 key 又是按位置分配的，
      // 于是选项本来就是 A/B/C 顺序时映射退化成恒等映射 —— 答案字母不跟着乱序走，
      // 导致开启"选项乱序"后正确答案指向别的选项（答对被判错）。副本里 o.key 此刻仍是原 key。
      const keyMap = {};
      newQ.options.forEach((o, idx) => {
        const newKey = String.fromCharCode(65 + idx);
        keyMap[o.key] = newKey;
        o.key = newKey;
      });
      // v2.7.6: 先把答案归一成「洗牌前的选项 key」，再逐个映射到新 key。
      // 原实现按字符拆分原答案再查表，只能处理单字母答案；判断题答案若还是
      // 「正确 / 错误」这类文字（老数据、手改过的 JSON），拆开后查不到任何 key，
      // 答案留在原地而选项已经打乱 —— 判分就会指向另一个选项。
      newQ.answer = getCorrectAnswers(q).map(k => keyMap[k]).filter(Boolean).join('');
      return newQ;
    });
  }

  startPracticeMode(questions, label, startIndex = 0) {
    this.currentQuestions = this.applyOptionShuffle(questions);
    const start = Number(startIndex);
    const maxIndex = Math.max(0, this.currentQuestions.length - 1);
    this.currentIndex = Math.min(Math.max(0, Number.isFinite(start) ? Math.trunc(start) : 0), maxIndex);
    this.mode = 'practice';
    this.modeLabel = label;
    if (this.currentBank) {
      this.answerRecords = this.plugin.getModeAnswerRecords(this.currentBank.id, this.getPracticeKey());
    } else {
      this.answerRecords = {};
    }
    this.resetQuestionState();
    this.render();
  }

  startReviewMode(questions) {
    this.currentQuestions = this.applyOptionShuffle(questions);
    this.currentIndex = 0;
    this.mode = 'review';
    this.modeLabel = '背题模式';
    this.hasAnswered = false;
    this.selectedOptions.clear();
    this.render();
  }

  // v2.5: 答题界面字体大小调节（0.7x ~ 2.2x，步进 0.1，设置中持久化）
  adjustFont(dir) {
    const cur = typeof this.plugin.settings.quizFontScale === 'number' ? this.plugin.settings.quizFontScale : 1;
    const next = Math.min(2.2, Math.max(0.7, Math.round((cur + dir * 0.1) * 10) / 10));
    this.plugin.settings.quizFontScale = next;
    this.plugin.saveDataDebounced();
    this.applyFontScale();
    new Notice(`字体大小 ${Math.round(next * 100)}%`, 1500);
  }

  /** v2.5: 将字体缩放应用到当前视图容器 */
  applyFontScale() {
    const scale = typeof this.plugin.settings.quizFontScale === 'number' ? this.plugin.settings.quizFontScale : 1;
    if (this.contentEl) this.contentEl.style.setProperty('--quiz-font-scale', String(scale));
  }

  saveCurrentProgress() {
    if (!this.currentBank) return;
    const key = progressKeyForLabel(this.modeLabel);
    if (!key) return;   // 背题模式 / 搜索结果等不记进度
    this.plugin.saveProgress(this.currentBank.id, key, this.currentIndex, this.currentQuestions.length);
  }

  /**
   * v2.7.4: 进入练习的统一入口（"可续做"的入口会跳到上次位置）。
   * 只在题量与上次完全一致时才续做 —— 题库增删过就安全地从第 1 题开始。
   */
  startPracticeWithProgress(questions, label) {
    const key = progressKeyForLabel(label);
    let index = 0;
    if (key && RESUMABLE_PROGRESS_KEYS.has(key) && this.currentBank) {
      index = this.plugin.getSavedProgress(this.currentBank.id, key, questions.length);
    }
    this.startPracticeMode(questions, label, index);
    if (index > 0) new Notice(`已从上次的第 ${index + 1} 题继续（⋯ 菜单可从头开始）`, 3000);
  }
}

/* ============================================================
   Section 14: 设置面板
   ============================================================ */

class QuizSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }
  // Obsidian 1.13+：addSettingTab 注册时调用 update()（兼容无参调用）
  update(containerEl) {
    this._render(containerEl || this.containerEl || document.body);
  }
  // Obsidian ≤1.12：display(containerEl)
  display(containerEl) {
    this._render(containerEl || this.containerEl || document.body);
  }
  _render(containerEl) {
    containerEl.empty();
    new Setting(containerEl)
      .setName('答对自动跳下一题')
      .setDesc('单选题/判断题答对后自动切换到下一题')
      .addToggle(toggle => {
        toggle.setValue(this.plugin.settings.autoNextOnCorrect);
        toggle.onChange(async (value) => { this.plugin.settings.autoNextOnCorrect = value; await this.plugin.saveDataDebounced(); });
      });
    new Setting(containerEl)
      .setName('选项乱序')
      .setDesc('每次进入题目时随机打乱选项顺序')
      .addToggle(toggle => {
        toggle.setValue(this.plugin.settings.shuffleOptions);
        toggle.onChange(async (value) => { this.plugin.settings.shuffleOptions = value; await this.plugin.saveDataDebounced(); });
      });
    new Setting(containerEl)
      .setName('答对自动移出错题集')
      .setDesc('答题正确时自动从错题集中移除该题（默认关闭，错题仅通过手动移除/清空）')
      .addToggle(toggle => {
        toggle.setValue(this.plugin.settings.autoRemoveWrongOnCorrect);
        toggle.onChange(async (value) => { this.plugin.settings.autoRemoveWrongOnCorrect = value; await this.plugin.saveDataDebounced(); });
      });
    new Setting(containerEl)
      .setName('相似题推荐阈值')
      .setDesc('相似度达到该阈值（0-1）的题目才会被推荐，值越高匹配越严格')
      .addSlider(slider => {
        slider.setLimits(0, 1, 0.05);
        slider.setValue(this.plugin.settings.similarityThreshold);
        slider.setDynamicTooltip();
        slider.onChange(async (value) => { this.plugin.settings.similarityThreshold = value; await this.plugin.saveDataDebounced(); });
      });
    containerEl.createEl('h3', { text: '数据管理' });
    new Setting(containerEl)
      .setName('导出数据')
      .setDesc('将所有题库和答题记录导出为 JSON 文件')
      .addButton(btn => {
        btn.setButtonText('导出');
        btn.onClick(async () => {
          const data = await this.plugin.exportAllData();
          const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `quiz-bank-backup-${new Date().toISOString().slice(0, 10)}.json`;
          a.click();
          URL.revokeObjectURL(url);
          new Notice('数据已导出');
        });
      });
    new Setting(containerEl)
      .setName('导入数据')
      .setDesc('从 JSON 备份文件恢复数据')
      .addButton(btn => {
        btn.setButtonText('导入');
        btn.onClick(() => {
          const input = document.createElement('input');
          input.type = 'file';
          input.accept = '.json';
          input.onchange = async (e) => {
            const file = e.target.files[0];
            if (!file) return;
            try {
              const text = await file.text();
              await this.plugin.importAllData(JSON.parse(text));
              new Notice('数据导入成功');
            } catch (err) { new Notice(`导入失败: ${err.message}`); }
          };
          input.click();
        });
      });
  }
}

/* ============================================================
   Section 15: Android 移动端 Modal 包装器
   ============================================================ */

class QuizMobileModal extends Modal {
  /**
   * Android 端用 Modal 替代 ItemView 运行题库刷题。
   *
   * v4 方案：不执行 new QuizView()（Obsidian 的 View 构造函数会访问
   * leaf.view.children 等 fakeLeaf 缺失属性 → Cannot read properties of
   * undefined）。改用 Object.create(QuizView.prototype) 跳过构造函数，
   * 手动补全实例字段；contentEl 在实例上从未定义过，普通赋值即生效，
   * 无任何 getter/描述符限制。全部 render 逻辑原样复用。
   */
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this._quizView = null;
  }

  onOpen() {
    const { contentEl, modalEl, titleEl } = this;

    // 0. 全屏样式通过 styles.css 的 .quiz-mobile-modal / .quiz-mobile-content 类实现（不再内联）
    modalEl.addClass('quiz-mobile-modal');
    contentEl.addClass('quiz-mobile-content');

    try {
      // 1. 跳过 View/ItemView 构造函数，直接以 QuizView.prototype 创建实例
      const view = Object.create(QuizView.prototype);

      // 通过 _initFields 统一初始化字段（与桌面 QuizView 同一来源，避免新增字段漏初始化）
      view._initFields(this.plugin);
      view.isModalHost = true;   // v2.7.4: 无真实 leaf，键盘焦点判断直接放行
      view.app = this.app;
      view.leaf = { app: this.app, view: { containerEl: document.createElement('div') } };

      // 2. ★ contentEl 直接指向 Modal 内容区（实例上无此属性，赋值即创建）
      view.contentEl = contentEl;
      view.containerEl = contentEl;
      this._quizView = view;

      // 3. 渲染题库 UI（异步，错误显示在弹窗内便于排查）
      view.render().catch((err) => {
        console.error('[QuizMobileModal] render error:', err);
        contentEl.empty();
        contentEl.createEl('div', {
          cls: 'quiz-empty-text',
          text: '渲染出错: ' + (err && err.message ? err.message : err)
        });
      });

      // 4. 键盘快捷键（A/B/C/D 选选项、方向键切换等）
      view.setupKeyboard();
    } catch (err) {
      console.error('[QuizMobileModal] open error:', err);
      contentEl.empty();
      contentEl.createEl('div', {
        cls: 'quiz-empty-text',
        text: '打开出错: ' + (err && err.message ? err.message : err)
      });
    }
  }

  onClose() {
    if (this._quizView) {
      this._quizView.cleanupKeyboard();
      this._quizView.removeContextMenu();
      // v2.7.6: Modal 宿主没有 View 卸载钩子，旋转监听要手动摘掉
      if (this._quizView._layoutResizeHandler) {
        window.removeEventListener('resize', this._quizView._layoutResizeHandler);
        this._quizView._layoutResizeHandler = null;
        this._quizView._layoutResizeBound = false;
      }
      this._quizView = null;
    }
    this.contentEl.empty();
  }
}

/* ============================================================
   Section 16: 主插件类
   ============================================================ */

class QuizBankPlugin extends Plugin {
  async onload() {
    await this.loadData();
    this.settings = Object.assign({
      autoNextOnCorrect: true,
      shuffleOptions: false,
      similarityThreshold: 0.3,
      autoRemoveWrongOnCorrect: false,
      quizFontScale: 1
    }, this.data?.settings);

    this.banks = {};
    this.userData = {};
    this.searchIndex = null;
    this._dirtyBanks = new Set();
    this._dirtyRecords = new Set();

    // 从独立文件加载所有题库和答题记录
    await this.loadAllBanksAndRecords();

    // 迁移：如果 data.json 中存在旧格式数据，迁移到独立文件
    if (this.data?.banks || this.data?.userData) {
      if (this.data.banks) {
        for (const [id, bank] of Object.entries(this.data.banks)) {
          if (!isValidBank(bank)) {
            console.warn(`跳过 data.json 中结构损坏的题库: ${id || '(未知)'}`);
            continue;
          }
          if (!this.banks[id]) {
            this.banks[id] = sanitizeBankQuestions(bank);
            await this.writeBankFile(id);
          }
        }
      }
      if (this.data.userData) {
        for (const [id, ud] of Object.entries(this.data.userData)) {
          if (!this.userData[id]) {
            this.userData[id] = ud;
            ud.bankId = id;
            await this.writeRecordsFile(id);
          }
        }
      }
      // 清除 data.json 中的旧数据，只保留 settings
      await super.saveData({ settings: this.settings });
    }

    this.saveDataDebounced = debounce(() => this.saveData(), 500);

    // 安卓端：使用 Modal 包装器代替 ItemView（安卓 ItemView 无法正常渲染）
    if (isAndroid()) {
      this.addCommand({ id: 'open', name: '打开题库刷题', callback: () => new QuizMobileModal(this.app, this).open() });
      this.addRibbonIcon('book-open', '题库刷题', () => new QuizMobileModal(this.app, this).open());
    } else {
      this.registerView(VIEW_TYPE_QUIZ, (leaf) => new QuizView(leaf, this));
      this.addCommand({ id: 'open', name: '打开题库刷题', callback: () => this.activateView() });
      this.addRibbonIcon('book-open', '题库刷题', () => this.activateView());
    }

    this.addSettingTab(new QuizSettingTab(this.app, this));
  }

  async onunload() {
    try { await this.saveData(); } catch (e) { console.error('Save on unload failed:', e); }
  }

  async activateView() {
    const { workspace } = this.app;
    let leaf = workspace.getLeavesOfType(VIEW_TYPE_QUIZ)[0];
    if (!leaf) {
      leaf = workspace.getLeaf('tab');
      await leaf.setViewState({ type: VIEW_TYPE_QUIZ, active: true });
    }
    workspace.revealLeaf(leaf);
  }

  async loadData() { this.data = await super.loadData() || {}; }

  async saveData() {
    // 保存设置到 data.json
    await super.saveData({ settings: this.settings });
    // 保存脏题库文件
    for (const bankId of this._dirtyBanks) {
      await this.writeBankFile(bankId);
    }
    this._dirtyBanks.clear();
    // 保存脏答题记录文件
    for (const bankId of this._dirtyRecords) {
      await this.writeRecordsFile(bankId);
    }
    this._dirtyRecords.clear();
  }

  /* ===== 文件路径与工具方法 ===== */

  get pluginDir() {
    return this.manifest.dir;
  }

  /** v2.6: 题库数据子目录（题库与答题记录统一存放，避免插件根目录混乱） */
  get dataDir() {
    return `${this.pluginDir}/data`;
  }

  /** 确保数据子目录存在（不存在则自动创建） */
  async ensureDataDir() {
    const adapter = this.app.vault.adapter;
    try {
      if (!(await adapter.exists(this.dataDir))) await adapter.mkdir(this.dataDir);
    } catch (e) {
      console.error('创建题库数据目录失败:', e);
    }
  }

  /** 清理文件名中的非法字符 */
  sanitizeFileName(name) {
    return (name || '').replace(/[\\/:*?"<>|]/g, '_');
  }

  /** 题库文件路径（基于题库 ID，保存在 data 子目录） */
  bankFilePath(bankId) {
    return `${this.dataDir}/bank_${bankId}.json`;
  }

  /** 答题记录文件路径（基于题库 ID，保存在 data 子目录） */
  recordsFilePath(bankId) {
    return `${this.dataDir}/bank_${bankId}_records.json`;
  }

  /** 标记题库为脏（需要保存） */
  markBankDirty(bankId) { this._dirtyBanks.add(bankId); }

  /** 标记答题记录为脏（需要保存） */
  markRecordsDirty(bankId) { this._dirtyRecords.add(bankId); }

  /** 写入题库文件 */
  async writeBankFile(bankId) {
    const bank = this.banks[bankId];
    if (!bank) return;
    await this.ensureDataDir();
    const adapter = this.app.vault.adapter;
    const path = this.bankFilePath(bank.id);
    await adapter.write(path, JSON.stringify(bank, null, 2));
  }

  /** 写入答题记录文件 */
  async writeRecordsFile(bankId) {
    const ud = this.userData[bankId];
    if (!ud) return;
    const bank = this.banks[bankId];
    if (!bank) return;
    ud.bankId = bankId;
    await this.ensureDataDir();
    const adapter = this.app.vault.adapter;
    const path = this.recordsFilePath(bank.id);
    await adapter.write(path, JSON.stringify(ud, null, 2));
  }

  /** 从独立文件加载所有题库和答题记录；自动迁移旧版（按题库名命名）文件到新版（按 ID 命名） */
  async loadAllBanksAndRecords() {
    const adapter = this.app.vault.adapter;
    const pluginDir = this.pluginDir;
    const dataDir = this.dataDir;
    try {
      await this.ensureDataDir();

      // v2.6 第 0 遍：把旧版本散落在插件根目录的题库/记录 JSON（含备份）迁移进 data 子目录
      const rootListing = await adapter.list(pluginDir);
      for (const filePath of rootListing.files) {
        const fileName = filePath.split(/[\\/]/).pop();
        if (!fileName.endsWith('.json') || fileName === 'data.json' || fileName === 'manifest.json') continue;
        const targetPath = `${dataDir}/${fileName}`;
        if (await adapter.exists(targetPath)) {
          console.warn(`数据目录已存在同名文件，保留数据目录版本: ${fileName}`);
          continue;
        }
        try {
          await adapter.write(targetPath, await adapter.read(filePath));
          await adapter.remove(filePath);
          console.log(`已迁移数据文件: ${fileName} -> data/${fileName}`);
        } catch (e) {
          console.error(`迁移数据文件失败 ${fileName}:`, e);
        }
      }

      const listing = await adapter.list(dataDir);
      const recordsFiles = [];
      const legacyBankFiles = [];

      // 第一遍：扫描文件，区分新格式 / 旧格式 / 记录文件
      for (const filePath of listing.files) {
        const fileName = filePath.split(/[\\/]/).pop();
        if (!fileName.endsWith('.json') || fileName === 'data.json' || fileName === 'manifest.json') continue;
        if (fileName.startsWith('backup_')) continue;
        if (fileName.endsWith('_records.json')) { recordsFiles.push(filePath); continue; }
        if (/^bank_[A-Za-z0-9]+\.json$/.test(fileName)) {
          try {
            const content = await adapter.read(filePath);
            const data = JSON.parse(content);
            // 校验结构：缺 id / name / questions 数组的视为损坏文件，跳过并告警
            if (isValidBank(data)) {
              // v2.6.1: 载入时规范化答案/选项 key（修复全角字母、「、」等导致的多选误判）
              // v2.7.4: 同时做结构清洗，损坏的单题不再让整个渲染/搜索流程抛错
              sanitizeBankQuestions(data);
              this.banks[data.id] = data;
            } else {
              console.warn(`跳过结构损坏的题库文件: ${fileName}（缺少 id/name/questions）`);
            }
          } catch (e) {
            console.error(`加载题库文件失败 ${filePath}:`, e);
          }
        } else {
          legacyBankFiles.push(filePath);
        }
      }

      // 第二遍：加载旧格式题库文件并迁移到 bank_<id>.json
      for (const filePath of legacyBankFiles) {
        const fileName = filePath.split(/[\\/]/).pop();
        try {
          const content = await adapter.read(filePath);
          const data = JSON.parse(content);
          if (!isValidBank(data)) {
            console.warn(`跳过结构损坏的题库文件: ${fileName}`);
            continue;
          }
          sanitizeBankQuestions(data);
          this.banks[data.id] = data;
          const newPath = this.bankFilePath(data.id);
          if (await adapter.exists(newPath)) {
            console.warn(`目标文件已存在，跳过迁移: ${fileName}`);
            continue;
          }
          try {
            await adapter.write(newPath, JSON.stringify(data, null, 2));
            await adapter.remove(filePath);
            console.log(`已迁移题库文件: ${fileName} -> bank_${data.id}.json`);
          } catch (e) {
            console.error(`迁移题库文件失败 ${fileName}:`, e);
          }
        } catch (e) {
          console.error(`加载题库文件失败 ${filePath}:`, e);
        }
      }

      // 第三遍：加载答题记录文件（兼容旧格式并迁移）
      for (const recordsPath of recordsFiles) {
        const fileName = recordsPath.split(/[\\/]/).pop();
        try {
          const content = await adapter.read(recordsPath);
          const data = JSON.parse(content);
          const bankId = data && data.bankId;
          if (bankId) {
            if (this.banks[bankId]) this.userData[bankId] = data;
            // 迁移旧命名 <name>_records.json -> bank_<id>_records.json
            if (!/^bank_[A-Za-z0-9]+_records\.json$/.test(fileName)) {
              const newPath = this.recordsFilePath(bankId);
              if (await adapter.exists(newPath)) {
                console.warn(`目标文件已存在，跳过迁移: ${fileName}`);
                continue;
              }
              try {
                await adapter.write(newPath, JSON.stringify(data, null, 2));
                await adapter.remove(recordsPath);
                console.log(`已迁移记录文件: ${fileName} -> bank_${bankId}_records.json`);
              } catch (e) {
                console.error(`迁移记录文件失败 ${fileName}:`, e);
              }
            }
          }
        } catch (e) {
          console.error(`加载记录文件失败 ${recordsPath}:`, e);
        }
      }
    } catch (e) {
      console.error('列出插件目录失败:', e);
    }
  }

  /** 清除所有题库和记录文件（用于导入恢复；仅清 data 子目录，备份文件保留） */
  async clearAllBankFiles() {
    const adapter = this.app.vault.adapter;
    try {
      if (!(await adapter.exists(this.dataDir))) return;
      const listing = await adapter.list(this.dataDir);
      for (const filePath of listing.files) {
        const fileName = filePath.split(/[\\/]/).pop();
        if (fileName.startsWith('backup_')) continue;
        if (fileName.endsWith('.json')) {
          await adapter.remove(filePath);
        }
      }
    } catch (e) {
      console.error('清除题库文件失败:', e);
    }
  }

  /* ===== 题库操作 ===== */

  async getAllBanks() {
    return Object.values(this.banks);
  }

  invalidateCache() { this.searchIndex = null; }

  /** 构建/获取搜索索引（预小写化，避免每次按键全库重复 toLowerCase；数据变更时由 invalidateCache 失效） */
  getSearchIndex() {
    if (!this.searchIndex) {
      this.searchIndex = {};
      for (const [id, bank] of Object.entries(this.banks)) {
        this.searchIndex[id] = bank.questions.map(q => ({
          q,
          textLower: String(q.text || '').toLowerCase(),
          optionsLower: (Array.isArray(q.options) ? q.options : []).map(o => String((o && o.text) || '').toLowerCase())
        }));
      }
    }
    return this.searchIndex;
  }

  async createBank(name, questions) {
    const id = generateId();
    const bank = { id, name, questions, createdAt: Date.now() };
    this.banks[id] = bank;
    this.getUserData(id); // 初始化空的用户数据
    await this.writeBankFile(id);
    await this.writeRecordsFile(id);
    this.invalidateCache();
    return bank;
  }



  async renameBank(bankId, newName) {
    const bank = this.banks[bankId];
    if (!bank) return;
    if (bank.name === newName) return;
    // 文件按 bankId 命名，改名无需移动文件，也避免重名误删他人数据
    bank.name = newName;
    this.markBankDirty(bankId);
    this.invalidateCache();
  }

  async deleteBank(bankId) {
    const bank = this.banks[bankId];
    if (!bank) return;
    const adapter = this.app.vault.adapter;
    const bankPath = this.bankFilePath(bank.id);
    const recordsPath = this.recordsFilePath(bank.id);

    // 同步删除题库文件和答题记录文件
    if (await adapter.exists(bankPath)) {
      await adapter.remove(bankPath);
    }
    if (await adapter.exists(recordsPath)) {
      await adapter.remove(recordsPath);
    }

    // 从内存中移除
    delete this.banks[bankId];
    if (this.userData[bankId]) delete this.userData[bankId];
    this._dirtyBanks.delete(bankId);
    this._dirtyRecords.delete(bankId);
    this.invalidateCache();
  }

  async updateQuestion(bankId, question) {
    const bank = this.banks[bankId];
    if (!bank) return;
    const idx = bank.questions.findIndex(q => q.id === question.id);
    if (idx !== -1) {
      bank.questions[idx] = question;
      this.markBankDirty(bankId);
      this.invalidateCache();
    }
  }

  async deleteQuestion(bankId, questionId) {
    const bank = this.banks[bankId];
    if (!bank) return;
    bank.questions = bank.questions.filter(q => q.id !== questionId);
    this.markBankDirty(bankId);
    this.invalidateCache();

    const ud = this.getUserData(bankId);
    if (ud.wrongQuestions) ud.wrongQuestions = ud.wrongQuestions.filter(id => id !== questionId);
    if (ud.favoriteQuestions) ud.favoriteQuestions = ud.favoriteQuestions.filter(id => id !== questionId);
    if (ud.answerRecords) delete ud.answerRecords[questionId];
    if (ud.modeRecords) { Object.values(ud.modeRecords).forEach(recs => { delete recs[questionId]; }); }
    this.markRecordsDirty(bankId);
  }

  /* ===== 用户数据操作 ===== */

  getUserData(bankId) {
    if (!this.userData[bankId] || typeof this.userData[bankId] !== 'object') {
      this.userData[bankId] = {
        bankId,
        wrongQuestions: [],
        favoriteQuestions: [],
        answerRecords: {},
        modeRecords: { practice: {}, wrong: {}, favorites: {}, group: {}, search: {} },
        progress: {}
      };
    }
    // v2.7.4: 逐字段兜底。导入的备份或旧版记录文件可能缺字段，
    // 原来只补 modeRecords，于是第一次答错时 addWrongQuestion 里的
    // ud.wrongQuestions.includes 会直接抛 TypeError，整个答题流程中断。
    const ud = this.userData[bankId];
    if (!Array.isArray(ud.wrongQuestions)) ud.wrongQuestions = [];
    if (!Array.isArray(ud.favoriteQuestions)) ud.favoriteQuestions = [];
    if (!ud.answerRecords || typeof ud.answerRecords !== 'object') ud.answerRecords = {};
    if (!ud.modeRecords || typeof ud.modeRecords !== 'object') ud.modeRecords = {};
    if (!ud.progress || typeof ud.progress !== 'object') ud.progress = {};
    if (!ud.bankId) ud.bankId = bankId;
    return ud;
  }

  getWrongQuestions(bankId) { return this.getUserData(bankId).wrongQuestions || []; }
  addWrongQuestion(bankId, qid) {
    const ud = this.getUserData(bankId);
    if (!ud.wrongQuestions.includes(qid)) ud.wrongQuestions.push(qid);
    this.markRecordsDirty(bankId);
  }
  removeWrongQuestion(bankId, qid) {
    const ud = this.getUserData(bankId);
    ud.wrongQuestions = ud.wrongQuestions.filter(id => id !== qid);
    this.markRecordsDirty(bankId);
  }
  clearWrongQuestions(bankId) {
    this.getUserData(bankId).wrongQuestions = [];
    this.markRecordsDirty(bankId);
  }

  getFavoriteQuestions(bankId) { return this.getUserData(bankId).favoriteQuestions || []; }
  addFavoriteQuestion(bankId, qid) {
    const ud = this.getUserData(bankId);
    if (!ud.favoriteQuestions.includes(qid)) ud.favoriteQuestions.push(qid);
    this.markRecordsDirty(bankId);
  }
  removeFavoriteQuestion(bankId, qid) {
    const ud = this.getUserData(bankId);
    ud.favoriteQuestions = ud.favoriteQuestions.filter(id => id !== qid);
    this.markRecordsDirty(bankId);
  }

  getAnswerRecords(bankId) { return this.getUserData(bankId).answerRecords || {}; }
  setAnswerRecord(bankId, qid, record) {
    this.getUserData(bankId).answerRecords[qid] = record;
    this.markRecordsDirty(bankId);
  }
  removeAnswerRecord(bankId, qid) {
    delete this.getUserData(bankId).answerRecords[qid];
    this.markRecordsDirty(bankId);
  }

  getModeAnswerRecords(bankId, modeKey) {
    const ud = this.getUserData(bankId);
    if (!ud.modeRecords[modeKey]) ud.modeRecords[modeKey] = {};
    return ud.modeRecords[modeKey];
  }
  setModeAnswerRecord(bankId, modeKey, qid, record) {
    const ud = this.getUserData(bankId);
    if (!ud.modeRecords[modeKey]) ud.modeRecords[modeKey] = {};
    ud.modeRecords[modeKey][qid] = record;
    this.markRecordsDirty(bankId);
  }

  /**
   * v2.7.4: 记录练习进度。同时记下题量 total —— 只有题量与上次一致时才允许续做，
   * 这样旧版本遗留的 progress 数据（键用的是视图 mode、语义混乱）不会被误用。
   */
  saveProgress(bankId, mode, index, total) {
    const ud = this.getUserData(bankId);
    ud.progress[mode] = { index, total: total || 0, timestamp: Date.now() };
    this.markRecordsDirty(bankId);
    this.saveDataDebounced();
  }

  /** v2.7.4: 读取可续做的进度位置（无记录 / 题量不一致 / 越界时返回 0） */
  getSavedProgress(bankId, mode, total) {
    const progress = this.getUserData(bankId).progress || {};
    const saved = progress[mode];
    if (!saved || typeof saved.index !== 'number') return 0;
    if (!saved.total || saved.total !== total) return 0;
    if (saved.index <= 0 || saved.index >= total) return 0;
    return saved.index;
  }

  /* ===== 导入导出 ===== */

  async exportAllData() {
    return {
      version: DATA_VERSION,
      exportedAt: new Date().toISOString(),
      settings: this.settings,
      banks: this.banks,
      userData: this.userData
    };
  }

  async importAllData(data) {
    if (!data || !data.banks || typeof data.banks !== 'object') throw new Error('无效的备份文件');

    // v2.7.4: 先校验再动手。结构损坏的题库直接剔除（而不是原样写回磁盘后
    // 在渲染/搜索时抛错），并且必须在清空现有数据之前完成，
    // 避免出现"旧数据清空了、新数据又导不进来"的最坏情况。
    const incomingBanks = {};
    let skippedBanks = 0;
    for (const [bankId, bank] of Object.entries(data.banks)) {
      if (!isValidBank(bank)) { skippedBanks++; continue; }
      incomingBanks[bankId] = sanitizeBankQuestions(bank);
    }
    if (Object.keys(incomingBanks).length === 0) throw new Error('备份文件中没有可用的题库');

    // 1. 先备份当前全部数据（防止清空后写入失败导致数据丢失）
    const backupFileName = `backup_${Date.now()}.json`;
    const backupPath = `${this.dataDir}/${backupFileName}`;
    await this.ensureDataDir();
    try {
      const backupData = {
        version: DATA_VERSION,
        exportedAt: new Date().toISOString(),
        settings: this.settings,
        banks: this.banks,
        userData: this.userData
      };
      await this.app.vault.adapter.write(backupPath, JSON.stringify(backupData, null, 2));
    } catch (e) {
      console.error('备份当前数据失败:', e);
      throw new Error('备份当前数据失败，已中止导入：' + e.message);
    }

    // 2. 清除现有题库和记录文件（备份文件不受影响）
    await this.clearAllBankFiles();

    // 3. 更新内存状态（只保留有对应题库的记录，避免产生永远读不到的孤儿记录）
    this.banks = incomingBanks;
    this.userData = {};
    if (data.userData && typeof data.userData === 'object') {
      for (const [bankId, ud] of Object.entries(data.userData)) {
        if (this.banks[bankId] && ud && typeof ud === 'object') this.userData[bankId] = ud;
      }
    }
    if (data.settings) this.settings = Object.assign(this.settings, data.settings);
    this.searchIndex = null;

    // 确保所有用户数据都有 bankId
    for (const [bankId, ud] of Object.entries(this.userData)) {
      ud.bankId = bankId;
    }

    // 4. 写入所有题库和记录文件
    for (const bankId of Object.keys(this.banks)) {
      await this.writeBankFile(bankId);
      if (this.userData[bankId]) {
        await this.writeRecordsFile(bankId);
      }
    }

    // 5. 保存设置
    await super.saveData({ settings: this.settings });
    new Notice(`原数据已自动备份为 ${backupFileName}（可随时从设置中导入恢复）`);
    if (skippedBanks > 0) new Notice(`已跳过 ${skippedBanks} 个结构损坏的题库`, 5000);
  }
}

module.exports = QuizBankPlugin;
