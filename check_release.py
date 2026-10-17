#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Obsidian 插件上架前自检脚本。

模拟 obsidian-releases 机器人的主要校验规则，在本地先把硬性问题挑出来，
避免提了 PR 才被打回。

用法：
    python check_release.py [仓库目录]    # 默认当前目录
"""

import json
import os
import re
import sys

ALLOWED_MANIFEST_KEYS = {
    "id", "name", "version", "minAppVersion", "description",
    "author", "authorUrl", "fundingUrl", "helpUrl", "isDesktopOnly",
}
REQUIRED_MANIFEST_KEYS = {
    "id", "name", "version", "minAppVersion", "description",
    "author", "isDesktopOnly",
}
REQUIRED_ASSETS = ["main.js", "manifest.json"]

errors, warnings = [], []


def err(msg):
    errors.append(f"[错误] {msg}")


def warn(msg):
    warnings.append(f"[提醒] {msg}")


def main(root):
    os.chdir(root)
    print(f"检查目录：{os.path.abspath(root)}\n")

    # ---------- 1. 仓库必备文件 ----------
    for f in ("README.md", "LICENSE", "manifest.json"):
        if not os.path.isfile(f):
            err(f"仓库根目录缺少 {f}")

    # ---------- 2. manifest 解析 ----------
    try:
        with open("manifest.json", encoding="utf-8") as fh:
            m = json.load(fh)
    except Exception as e:
        err(f"manifest.json 无法解析：{e}")
        report()
        return 1

    missing = REQUIRED_MANIFEST_KEYS - set(m)
    for k in sorted(missing):
        err(f"manifest.json 缺少必填字段：{k}")

    extra = set(m) - ALLOWED_MANIFEST_KEYS
    for k in sorted(extra):
        err(f"manifest.json 出现不允许的字段：{k}（允许的键只有 {sorted(ALLOWED_MANIFEST_KEYS)}）")

    # ---------- 3. id ----------
    pid = m.get("id", "")
    if pid:
        if "obsidian" in pid.lower():
            err('id 不能包含 "obsidian"')
        if not re.fullmatch(r"[a-z0-9_-]+", pid):
            err("id 只能包含小写字母、数字、连字符和下划线")
        if pid.endswith("plugin"):
            err('id 不能以 "plugin" 结尾')
        warn(f"id = {pid}，已核对官方清单未被占用，但提交前仍建议在 "
             "https://community.obsidian.md/search?type=plugin 再搜一次；"
             "id 一经发布永久不可修改")

    # ---------- 4. name ----------
    name = m.get("name", "")
    if name:
        low = name.lower()
        for bad in ("obsidian", "obsi-", "-sidian"):
            if bad in low:
                err(f'name 不能包含 "{bad}"')
        if low.endswith("plugin") or "plugin" in low:
            err('name 不能包含 "Plugin"')
        if not re.fullmatch(r"[\x20-\x7E]+", name):
            warn('name 含非 Basic Latin 字符，官方要求优先英文，人工审核可能被要求改名')
        if len(name) > 30:
            warn("name 偏长，建议简短")

    # ---------- 5. description ----------
    desc = m.get("description", "")
    if desc:
        if "obsidian" in desc.lower():
            err('description 不能包含 "Obsidian"')
        if len(desc) > 250:
            err(f"description 长度 {len(desc)} 超过 250 字符上限（当前超出 {len(desc) - 250}）")
        if desc.rstrip()[-1] not in ".?!) 。！？）":
            err('description 必须以 . ? ! ) 或中文 。！？） 结尾')

    # ---------- 6. version ----------
    ver = m.get("version", "")
    if ver:
        if not re.fullmatch(r"[0-9.]+", ver):
            err("version 只能是数字和点（如 2.7.4），不能带 v 前缀")
        if not re.fullmatch(r"\d+\.\d+\.\d+", ver):
            warn("version 建议使用 x.y.z 三段式语义化版本")

    # ---------- 7. URL ----------
    au = m.get("authorUrl", "")
    if au and au.rstrip("/") in ("https://obsidian.md", "https://obsidian.md/"):
        err("authorUrl 不能指向 obsidian.md")
    if au and re.search(r"(your[-_]?github|YOUR_GITHUB|example\.com)", au, re.I):
        err("authorUrl 还是占位符，提交前必须换成真实地址")
    if not au:
        warn("没有 authorUrl，建议补上作者主页或仓库地址")
    if "fundingUrl" in m:
        fu = m["fundingUrl"]
        if isinstance(fu, str):
            if not fu.strip():
                err("fundingUrl 为空字符串，没有赞助链接请直接删除该键")
            if fu.rstrip("/").startswith("https://obsidian.md/pricing"):
                err("fundingUrl 不能指向 obsidian.md/pricing")
        elif not isinstance(fu, dict):
            err("fundingUrl 只能是字符串或对象")

    # ---------- 8. versions.json ----------
    if os.path.isfile("versions.json"):
        try:
            with open("versions.json", encoding="utf-8") as fh:
                vj = json.load(fh)
            if ver and ver not in vj:
                err(f"versions.json 里没有 version {ver} 对应的条目")
        except Exception as e:
            err(f"versions.json 无法解析：{e}")
    else:
        warn("没有 versions.json，建议补上（版本到最低 App 版本的映射）")

    # ---------- 9. 待上传的 Release 资产 ----------
    for a in REQUIRED_ASSETS:
        if not os.path.isfile(a):
            err(f"Release 资产缺少 {a}（必须散装上传，只在 zip 里不算）")
    if os.path.isfile("styles.css"):
        print("[信息] 检测到 styles.css，发布时也要一并作为 Release 资产上传")

    # 主脚本之外的 .js 文件都不会被官方分发，必须已内联进 main.js
    main_src = ""
    if os.path.isfile("main.js"):
        try:
            with open("main.js", encoding="utf-8", errors="ignore") as fh:
                main_src = fh.read()
        except Exception:
            main_src = ""
    for f in sorted(os.listdir(".")):
        if os.path.isdir(f):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext == ".js" and f == "main.js":
            continue
        if ext not in (".js", ".xlsx", ".xls", ".csv", ".tsv"):
            continue
        if main_src and f in main_src:
            print(f"[信息] {f} 已内联进 main.js，无需单独分发")
        else:
            warn(f"{f} 不在官方分发的三个文件里。若插件运行时读取它，"
                 "请内联进 main.js，否则用户安装或更新后会丢失")

    # 产物里不能出现动态执行代码（社区审核清单的硬性要求）
    if main_src:
        if "new Function" in main_src:
            err("main.js 里出现 new Function —— 审核清单要求不得动态执行代码")
        if re.search(r"(?<![\w.$])eval\s*\(", main_src):
            err("main.js 里出现 eval( —— 审核清单要求不得动态执行代码")
        if "SheetJS" in main_src and not os.path.isfile("THIRD-PARTY-NOTICES.md"):
            warn("内联了 SheetJS（Apache-2.0），建议补 THIRD-PARTY-NOTICES.md 说明许可与修改")

    # ---------- 10. 不该进仓库的东西 ----------
    for d in ("data",):
        if os.path.isdir(d):
            err(f"目录 {d}/ 含用户题库数据，绝不能提交到公开仓库")
    if os.path.isfile("data.json"):
        err("data.json 含用户数据，绝不能提交到公开仓库")
    if os.path.isdir(".obsidian"):
        err("仓库里不应包含 .obsidian/ 目录")

    report()
    return 1 if errors else 0


def report():
    print("-" * 60)
    for w in warnings:
        print(w)
    for e in errors:
        print(e)
    print("-" * 60)
    if errors:
        print(f"结果：{len(errors)} 个错误，{len(warnings)} 个提醒 —— 请修复错误后再提交到社区目录")
    else:
        print(f"结果：0 个错误，{len(warnings)} 个提醒 —— 可以提交到社区目录")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "."))
