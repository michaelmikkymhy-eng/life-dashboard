"""
一键把 Obsidian 里所有 skill 软链接到指定 AI 工具的 skills 目录。

跨平台：Windows 优先 symlink，无权限降级 junction；macOS/Linux 用 symlink。

用法：
    python install_all.py                  # 全部装到 Trae（~/.trae-cn/skills）
    python install_all.py --tool claude    # 全部装到 Claude Code（~/.claude/skills）
    python install_all.py --dir <目录>     # 全部装到任意自定义目录
    python install_all.py --list           # 列出已识别的工具目录

源在 Obsidian（本目录下的各中文子目录），工具侧只是软链接，改源即全工具生效。
换机器 / 换工具后跑一遍即可。
"""
import os
import sys
import re
import shutil
import subprocess
import argparse

SKILLS_DIR = os.path.dirname(os.path.abspath(__file__))

# 已知 AI 工具的 skills 目录（按需扩展）
TOOLS = {
    "trae":   os.path.join(os.path.expanduser("~"), ".trae-cn", "skills"),
    "claude": os.path.join(os.path.expanduser("~"), ".claude", "skills"),
    "cursor": os.path.join(os.path.expanduser("~"), ".cursor", "skills"),
}


def read_name(skill_dir):
    """从 SKILL.md 的 frontmatter 读 name（英文标识）。"""
    p = os.path.join(skill_dir, "SKILL.md")
    with open(p, encoding="utf-8") as f:
        head = f.read(2048)
    m = re.search(r'^name:\s*["\']?(.+?)["\']?\s*$', head, re.M)
    return m.group(1).strip() if m else os.path.basename(skill_dir)


def is_link(path):
    """判断是否 symlink 或（Windows）junction。"""
    if os.path.islink(path):
        return True
    if sys.platform == "win32" and os.path.isdir(path):
        import ctypes
        attr = ctypes.windll.kernel32.GetFileAttributesW(path)
        return attr != -1 and bool(attr & 0x400)  # FILE_ATTRIBUTE_REPARSE_POINT
    return False


def realpath(p):
    return os.path.realpath(p)


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


def install_one(skill_dir, target_dir):
    name = read_name(skill_dir)
    link = os.path.join(target_dir, name)
    os.makedirs(target_dir, exist_ok=True)
    if is_link(link) and realpath(link) == realpath(skill_dir):
        print("[跳过] {0} 已链接且指向正确".format(name))
        return
    if os.path.lexists(link):
        remove(link)
    kind = make_link(skill_dir, link)
    print("[完成] {0} -> {1} ({2})".format(name, link, kind))
    if kind == "junction":
        print("[提示] junction 仅本机可用，跨平台请开启开发者模式后重跑以改用 symlink")


def main():
    ap = argparse.ArgumentParser(description="一键软链接所有 skill 到 AI 工具")
    ap.add_argument("--tool", choices=sorted(TOOLS), help="目标工具名")
    ap.add_argument("--dir", help="自定义 skills 目录")
    ap.add_argument("--list", action="store_true", help="列出已识别工具目录")
    a = ap.parse_args()
    if a.list:
        for k, v in TOOLS.items():
            print("{0:8} {1}".format(k, v))
        return
    target = os.path.expanduser(a.dir or (TOOLS[a.tool] if a.tool else TOOLS["trae"]))
    skills = [d for d in os.listdir(SKILLS_DIR)
              if os.path.isdir(os.path.join(SKILLS_DIR, d))
              and os.path.exists(os.path.join(SKILLS_DIR, d, "SKILL.md"))]
    if not skills:
        print("[无] 未找到任何含 SKILL.md 的 skill 子目录")
        return
    for d in sorted(skills):
        install_one(os.path.join(SKILLS_DIR, d), target)


if __name__ == "__main__":
    main()
