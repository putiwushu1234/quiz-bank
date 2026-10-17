#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把外部资源内联进 main.js，让插件变成完全自包含的单文件。

背景：Obsidian 安装 / 更新插件时只分发 main.js、manifest.json、styles.css 三个文件，
同目录的其他文件（xlsx.full.min.js、题库导入模板.xlsx）不会随插件安装，
更新后也不保证保留。因此把它们编译进 main.js，相关功能才不会在用户端失效。

处理两类资源：
  1. SheetJS 库 —— 用一层普通的函数作用域包装（等价于打包器生成的 CommonJS
     包装）把整份源码编译进 main.js，不使用 eval / new Function。
  2. 题库导入模板 —— 转 base64 常量，downloadTemplate() 读不到文件时用它兜底。

用法：
    python tools/inline_xlsx.py            # 默认读 tools/main.src.js，写 main.js
    python tools/inline_xlsx.py --check    # 只校验产物，不重新生成
"""

import argparse
import base64
import hashlib
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

SRC_MAIN = os.path.join(HERE, "main.src.js")
XLSX_LIB = os.path.join(ROOT, "xlsx.full.min.js")
TEMPLATE = os.path.join(ROOT, "题库导入模板.xlsx")
OUT_MAIN = os.path.join(ROOT, "main.js")

CONST_NAME = "loadInlineXLSX"
CONST_TPL = "TEMPLATE_XLSX_BASE64"

# tryRequireXLSX 里 require 候选全部失败后的插入点
ANCHOR = """  for (const name of candidates) {
    const found = probe(name);
    if (found) { XLSX = found; return XLSX; }
  }
  return null;
}"""

PATCH = """  for (const name of candidates) {
    const found = probe(name);
    if (found) { XLSX = found; return XLSX; }
  }

  // v2.7.6: 内联的 SheetJS（定义在文件末尾的 loadInlineXLSX）。
  // 它是一个普通的模块作用域包装函数，源码在构建时就已固定，
  // 运行时只是调用一次，不含任何动态代码求值。
  try {
    const found = typeof loadInlineXLSX === 'function' ? loadInlineXLSX() : null;
    if (isUsableXLSX(found)) { XLSX = found; return XLSX; }
  } catch (e) {
    console.warn('[QuizBank] 内联 XLSX 执行失败：', e && e.message);
  }
  return null;
}"""

# downloadTemplate() 读不到模板文件时的兜底插入点
ANCHOR_TPL = """    if (!buffer) {
      new Notice('未找到模板文件（预期位置：插件目录/题库导入模板.xlsx），请重新安装插件', 6000);
      return;
    }"""

PATCH_TPL = """    // v2.7.4: 模板同样不在官方分发的三个文件里，用 base64 内联副本兜底
    // （常量见文件末尾 TEMPLATE_XLSX_BASE64）。
    if (!buffer && typeof %(c)s === 'string' && %(c)s) {
      try {
        const bin = atob(%(c)s);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        buffer = arr.buffer;
        used = '内置模板';
      } catch (e) {
        console.warn('[QuizBank] 内置模板解码失败：', e && e.message);
      }
    }
    if (!buffer) {
      new Notice('未找到模板文件（预期位置：插件目录/题库导入模板.xlsx），请重新安装插件', 6000);
      return;
    }""" % {"c": CONST_TPL}

TAIL_COMMENT = """

/* ============================================================
   v2.7.6: 内联的 SheetJS 0.20.3 源码
   ------------------------------------------------------------
   Obsidian 只分发 main.js / manifest.json / styles.css，独立的
   xlsx.full.min.js 不会随插件安装，更新后也不保证保留，因此把整个库
   编译进 main.js，让插件保持单文件自包含。

   下面是一个普通的模块作用域包装（与打包器生成的 CommonJS 包装等价）：
   函数体在构建时就已确定，运行时只是被调用一次，不使用 eval，
   也不做任何动态代码求值。

   SheetJS (C) SheetJS LLC, Apache-2.0
   源码字节数：%d，sha256 前 12 位：%s
   ============================================================ */
"""


def read(path):
    with open(path, "r", encoding="utf-8", newline="") as fh:
        return fh.read()


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)


def build():
    if not os.path.isfile(SRC_MAIN):
        print(f"缺少源入口：{SRC_MAIN}")
        return 1
    if not os.path.isfile(XLSX_LIB):
        print(f"缺少库文件：{XLSX_LIB}")
        return 1

    main_src = read(SRC_MAIN)
    lib_src = read(XLSX_LIB)

    if ANCHOR not in main_src:
        print("未找到 tryRequireXLSX 的插入锚点，源文件结构可能与预期不符")
        return 1

    sha = hashlib.sha256(lib_src.encode("utf-8")).hexdigest()[:12]
    tail = TAIL_COMMENT % (len(lib_src.encode("utf-8")), sha)

    # 用普通函数作用域包住库源码：module / exports 作为局部变量传入，
    # 库的 UMD 分支因此把实例挂到我们的 module.exports 上，
    # 既不会污染 main.js 自身的 module.exports，也不需要动态执行。
    tail += (
        "function %s() {\n"
        "  const module = { exports: {} };\n"
        "  const exports = module.exports;\n"
        % CONST_NAME
    )
    tail += lib_src.rstrip("\n") + "\n"
    tail += "  return module.exports;\n}\n"

    out = main_src.replace(ANCHOR, PATCH, 1)

    tpl_note = "未内联（仓库里没有模板文件）"
    if os.path.isfile(TEMPLATE):
        if ANCHOR_TPL not in out:
            print("未找到 downloadTemplate 的插入锚点，模板未内联")
            return 1
        with open(TEMPLATE, "rb") as fh:
            raw = fh.read()
        b64 = base64.b64encode(raw).decode("ascii")
        out = out.replace(ANCHOR_TPL, PATCH_TPL, 1)
        tail += "\n/* 内置的题库导入模板，base64，原始字节数：%d */\nconst %s = \"%s\";\n" % (
            len(raw), CONST_TPL, b64
        )
        tpl_note = f"{len(raw) / 1024:.0f} KB → base64 {len(b64) / 1024:.0f} KB"

    out = out.rstrip("\n") + "\n" + tail

    write(OUT_MAIN, out)
    print(f"已生成 {OUT_MAIN}")
    print(f"  原 main.js : {len(main_src.encode('utf-8')) / 1024:.0f} KB")
    print(f"  内联库     : {len(lib_src.encode('utf-8')) / 1024:.0f} KB")
    print(f"  内联模板   : {tpl_note}")
    print(f"  产物       : {len(out.encode('utf-8')) / 1024:.0f} KB")
    return 0


def check():
    if not os.path.isfile(OUT_MAIN):
        print(f"产物不存在：{OUT_MAIN}")
        return 1
    out = read(OUT_MAIN)
    ok = True

    if out.count(f"function {CONST_NAME}(") != 1:
        print("内联加载函数缺失或重复，可能重复构建")
        ok = False
    if PATCH.split("for (const name of candidates)")[1].split("return null;")[0] not in out:
        print("tryRequireXLSX 里的内联分支缺失")
        ok = False
    if "open-quiz-bank" in out:
        print("仍然存在含插件 ID 的旧命令 ID")
        ok = False
    # 社区审核清单明确要求不得出现 eval / new Function
    if "new Function" in out:
        print("产物中出现 new Function —— 会被审核判定为动态执行代码")
        ok = False
    if re.search(r"(?<![\w.$])eval\s*\(", out):
        print("产物中出现 eval( —— 会被审核判定为动态执行代码")
        ok = False
    if re.search(r"\bmodule\.exports\s*=\s*XLSX", out):
        print("内联库可能污染了模块导出")
        ok = False
    if os.path.isfile(TEMPLATE):
        if out.count(f"const {CONST_TPL} = ") != 1:
            print("内置模板常量缺失或重复")
            ok = False
        if "内置模板" not in out:
            print("downloadTemplate 里的内置模板兜底分支缺失")
            ok = False

    print("校验通过" if ok else "校验失败")
    return 0 if ok else 1


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只校验产物")
    args = ap.parse_args()
    sys.exit(check() if args.check else build())
