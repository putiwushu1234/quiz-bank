# Quiz Bank

Practise quizzes inside Obsidian. Import question banks from CSV, TSV or Excel, drill single-choice, multiple-choice and true/false questions, and get wrong answers, favourites and progress tracked locally.

在 Obsidian 里直接刷题。导入 CSV / TSV / Excel 题库，做单选、多选、判断题，自动收集错题，智能组卷，复习进度全部存在本地。

## Installation / 安装

### Community plugins（after approval）

Settings → Community plugins → Browse, search for `Quiz Bank`, install and enable.

设置 → 第三方插件 → 浏览，搜索 `Quiz Bank`，安装并启用。

### BRAT（before approval / beta）

1. Install and enable [BRAT](https://github.com/TfTHacker/obsidian42-brat).
2. BRAT settings → Add beta plugin → `putiwushu1234/quiz-bank`.
3. Enable **Quiz Bank** in the plugin list.

1. 安装并启用 BRAT 插件。
2. BRAT 设置 → Add beta plugin → 填入 `putiwushu1234/quiz-bank`。
3. 在插件列表里启用 Quiz Bank。

### Manual / 手动安装

Download `main.js`, `manifest.json` and `styles.css` from [Releases](https://github.com/putiwushu1234/quiz-bank/releases) into `<vault>/.obsidian/plugins/quiz-bank/`, then restart Obsidian and enable the plugin. The Excel parser is bundled into `main.js`, so no extra files are needed.

从 Releases 下载 `main.js`、`manifest.json`、`styles.css` 三个文件，放进 `<你的库>/.obsidian/plugins/quiz-bank/`，重启 Obsidian 后启用。Excel 解析库已经编译进 `main.js`，不需要额外文件。

## Features / 功能

- **Import**: CSV, TSV and Excel (`.xls` / `.xlsx`) question banks, with a built-in template
- **Types**: single choice, multiple choice, true/false
- **Modes**: sequential, random, review (answers shown), chapter practice
- **Wrong answers & favourites**: collected automatically, editable by hand
- **Smart paper generation**: draw by chapter, type and count; multiple-choice can be filtered by "all / missing N options", true/false by "correct / wrong", each with its own exact count
- **Answer sheet**: see the whole paper and per-question status at any time
- **Similar questions**: find near-duplicates of the current question
- **Search & edit**: search inside the plugin, edit questions in place
- **Copy as callout**: turn a question into a collapsible callout for your notes
- **Resume**: reopen a practice and continue where you left off

- **题库导入**：支持 CSV、TSV、Excel（`.xls` / `.xlsx`）三种格式，内置模板
- **题型**：单选题、多选题、判断题
- **练习模式**：顺序练习、随机练习、背题模式（直接看答案）、章节练习
- **错题集与收藏**：自动归集做错的题，支持手动收藏
- **智能组卷**：按章节、题型、数量随机抽题生成试卷。多选题可按「全选 / 少选 N 项」筛选，判断题可按「✓ 正确 / ✕ 错误」筛选，两类都能单独指定精确抽取数量
- **答题卡**：答题过程中随时查看整体作答情况与统计
- **相似题推荐**：根据当前题目推荐相近题目
- **题目搜索与编辑**：在插件内搜索题目并直接修改
- **一键复制为 Callout**：把题目转成 Callout 格式粘贴到笔记里
- **继续上次进度**：关闭后再次打开可回到上次的位置

## Usage / 使用

1. Click the **book icon** in the ribbon, or run `Quiz Bank: 打开题库刷题` from the command palette.
2. Import a question bank — an Excel/CSV/TSV file, or use **下载模板** to get the template first.
3. Pick a practice mode and start answering.

1. 点击左侧活动栏的**书本图标**（或在命令面板执行 `Quiz Bank: 打开题库刷题`）。
2. 在插件内导入题库，不清楚格式时先点「⬇️ 下载 Excel 模板」拿一份模板。
3. 选择练习模式开始答题。

### 题库列格式

| 列名 | 说明 |
| --- | --- |
| 题干（或 题目） | 题目正文 |
| 选项A / 选项B … | 各选项分列；也可以只用一列「选项」，按字母标记切分 |
| 答案 | 单选 `A`，多选 `AC` 或 `A,C`，判断题兼容 `对/错`、`√/×`、`Y/N` |
| 题型 | 可选，缺省按选项与答案推断 |
| 解析 | 可选，解析归位到对应题目 |
| 章节 | 可选，用于章节练习与组卷 |

Excel 中的日期按显示文本读取，不会出现序列号错位。其余有表头的列会作为附加字段一起保存。

## Settings / 设置

设置 → 第三方插件 → Quiz Bank：

- **答对自动跳下一题**
- **选项乱序**（每次进入题目打乱选项顺序；判断题的「正确 / 错误」固定不打乱）
- **答对自动移出错题集**
- **相似题推荐阈值**
- **数据管理**：导出 / 导入 JSON 备份

答题界面右上角还有 `A-` / `A+` 调节字号；`⋯` 菜单里有「从头开始 / 复制题目 / 编辑题目 / 答题卡 / 相似题 / 删除此题」。

## Wide screens / 宽屏

The page fills the whole pane, and every block on a page shares the same left and right edges — the nav bar lines up with the content below it, on every view. The question page is the one exception: it narrows as a whole (nav, progress bar, question, explanation and action bar together) to a 1080px reading column, because question text longer than that is hard to read. The bank list is always **two equal columns**, whatever the window width.

页面宽度跟随编辑区，同一页面里所有区块（导航条、列表、章节树、搜索框、统计卡、底部浮动栏、空状态）共用同一左右边界，不会出现「导航比内容窄」的错位。答题页整体收窄到 1080px 阅读宽度——导航、进度条、题干、解析、操作栏一起收窄，所以该页内部同样对齐。题库列表在任何窗口宽度下都保持**两列等宽**，不会随屏幕变宽而变成三列、四列。

想整体更宽或更窄，改 `styles.css` 里的两个变量即可：`--quiz-read-width`（答题页阅读宽度）、`--quiz-page-width`（其余页面的内容宽度）。列表列数只在第 26 节一处定义，改那一处即可。

## Phones & tablets / 手机与平板

On phones the layout switches to a single column: the nav sticks to the top, the action bar sticks to the bottom and spans the full width of the screen, and every tap target is at least 44px. Tablets (≥769px) use the wide layout, and on touch devices (`hover: none`) the tap targets are raised to the same 44px, so a tablet is never harder to tap than a phone. Rotating the device re-evaluates the tablet layout without a reload.

手机端为单列布局：导航吸顶、底部操作栏吸底并横向铺满屏幕，点击热区不小于 44px。平板（≥769px）使用宽屏布局；触控平板（`hover: none`）会把点击热区补齐到与手机相同的 44px —— 平板不会比手机更难按。旋转屏幕后会重新判断平板布局，无需重开插件。

The only width rule that matters on a phone is the container's own padding; all the wide-screen limits above are scoped to `min-width: 769px`, so they never interfere with the full-bleed bottom bar.

手机端只受容器内边距影响，上面那些宽屏限宽规则全部限定在 `min-width: 769px`，不会干扰吸底操作栏的通栏效果。

## Data / 数据存放

All banks, answer records and favourites are plain JSON files under `<vault>/.obsidian/plugins/quiz-bank/data/`. Nothing is uploaded, and the plugin contains no telemetry and makes no network requests.

所有题库、答题记录、收藏都存在 `<你的库>/.obsidian/plugins/quiz-bank/data/` 下，纯本地文件，不上传任何服务器，插件不包含遥测或统计上报，也不发起任何网络请求。

> 注意：数据放在插件目录内，**卸载插件（删除插件文件夹）会一并删除题库**。重要题库请在设置里用「导出数据」留一份 JSON 备份。

## Development / 开发

The build source is `tools/main.src.js`; `main.js` is generated from it.

```bash
python tools/inline_xlsx.py          # 生成根目录的 main.js
python tools/inline_xlsx.py --check  # 校验产物（含「无 eval / Function 构造器」检查）
node tools/test_inline.js            # 验证内联的 SheetJS 能解析 Excel
node tools/audit_core.js             # 核心逻辑（解析 / 判分 / 乱序）边界用例
python check_release.py              # 上架自检
```

发布新版本：改 `manifest.json` 的 `version`，`versions.json` 追加一行，然后推同名 tag（如 `2.7.6`，**不要带 `v`**），GitHub Actions 会自动构建并创建 Release。

## Third-party components / 第三方组件

- [SheetJS](https://sheetjs.com/) 0.20.3 — Apache-2.0，见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)

Obsidian 只分发 `main.js`、`manifest.json`、`styles.css` 三个文件，所以 SheetJS 在构建时被编译进 `main.js`（见 `tools/inline_xlsx.py`），插件保持单文件自包含。升级 SheetJS 时替换仓库根目录的 `xlsx.full.min.js`，然后重跑构建即可。

## License

[MIT](LICENSE) © 灰辉故事绘（B站）
