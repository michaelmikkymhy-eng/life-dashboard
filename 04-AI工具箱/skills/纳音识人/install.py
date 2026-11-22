"""
把本 skill 软链接到指定 AI 工具的 skills 目录，让该工具能加载调用。

跨平台：Windows 优先 symlink，无权限降级 junction；macOS/Linux 用 symlink。

用法：
    python install.py                  # 默认装到 Trae（~/.trae-cn/skills）
    python install.py --tool claude    # 装到 Claude Code（~/.claude/skills）
    python install.py --dir <目录>     # 装到任意自定义 skills 目录
    python install.py --list           # 列出已识别的工具目录

源始终在本目录（Obsidian 内），工具侧只是软链接，改源即全工具自动生效。
换机器 / 换工具后跑一遍即可。
"""
import os
import sys
import re
import shutil
import subprocess
import argparse

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))

# 已知 AI 工具的 skills 目录（按需扩展）
TOOLS = {
    "trae":   os.path.join(os.path.expanduser("~"), ".trae-cn", "skills"),
    "claude": os.path.join(os.path.expanduser("~"), ".claude", "skills"),
    "cursor": os.path.join(os.path.expanduser("~"), ".cursor", "skills"),
}


def read_skill_name():
    """从 SKILL.md 的 frontmatter 读 name（英文标识，供工具调用）。"""
    p = os.path.join(SKILL_DIR, "SKILL.md")
    with open(p, encoding="utf-8") as f:
        head = f.read(2048)
    m = re.search(r'^name:\s*["\']?(.+?)["\']?\s*$', head, re.M)
    return m.group(1).strip() if m else os.path.basename(SKILL_DIR)


def is_link(path):
    """判断是否 symlink 或（Windows）junction。"""
    if os.path.islink(path):
        return True
    if sys.platform == "win32" and os.path.isdir(path):
        import ctypes
        attr = ctypes.windll.kernel32.GetFileAttributesW(path)
        return attr != -1 and bool(attr & 0x400)  # FILE_ATTRIBUTE_REPARSE_POINT
    return False


def realpath(path):
    return os.path.realpath(path)


def remove(path):
    if is_link(path):
        if sys.platform == "win32":
            subprocess.run(["cmd", "/c", "rmdir", path], check=True)
        else:
            os.unlink(path)
    elif os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.unlink(path)


def make_link(src, dst):
    if sys.platform == "win32":
        try:
            os.symlink(src, dst, target_is_directory=True)
            return "symlink"
        except OSError:
            subprocess.run(["cmd", "/c", "mklink", "/J", dst, src],
                           check=True, capture_output=True)
            return "junction"
    else:
        os.symlink(src, dst)
        return "symlink"


def install(target_dir):
    name = read_skill_name()
    link = os.path.join(target_dir, name)
    os.makedirs(target_dir, exist_ok=True)
    if is_link(link) and realpath(link) == realpath(SKILL_DIR):
        print("[跳过] {0} 已链接且指向正确".format(name))
        return
    if os.path.lexists(link):
        remove(link)
    kind = make_link(SKILL_DIR, link)
    print("[完成] {0} -> {1} ({2})".format(name, link, kind))
    if kind == "junction":
        print("[提示] junction 仅本机可用，跨平台请开启开发者模式后重跑以改用 symlink")


def main():
    ap = argparse.ArgumentParser(description="软链接本 skill 到 AI 工具的 skills 目录")
    ap.add_argument("--tool", choices=sorted(TOOLS), help="目标工具名")
    ap.add_argument("--dir", help="自定义 skills 目录")
    ap.add_argument("--list", action="store_true", help="列出已识别工具目录")
    a = ap.parse_args()
    if a.list:
        for k, v in TOOLS.items():
            print("{0:8} {1}".format(k, v))
        return
    target = a.dir or (TOOLS[a.tool] if a.tool else TOOLS["trae"])
    install(os.path.expanduser(target))


if __name__ == "__main__":
    main()
