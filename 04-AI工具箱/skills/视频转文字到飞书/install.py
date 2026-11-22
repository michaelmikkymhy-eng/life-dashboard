"""
把本 skill 链接到 Trae 的 skills 目录，让 Trae 能加载。

跨平台：Windows 优先 symlink，没权限时降级 junction；Mac/Linux 用 symlink。
换机器后跑一遍即可：
    python install.py
"""
import os
import sys
import subprocess
import shutil

SKILL_NAME = "video-transcript"  # 必须等于 SKILL.md frontmatter 里的 name
SKILL_SRC = os.path.dirname(os.path.abspath(__file__))
TRA_SKILLS = os.path.join(os.path.expanduser("~"), ".trae-cn", "skills")
LINK_PATH = os.path.join(TRA_SKILLS, SKILL_NAME)


def already_ok():
    if not os.path.islink(LINK_PATH) and not _is_junction(LINK_PATH):
        return False
    return os.path.realpath(LINK_PATH) == os.path.realpath(SKILL_SRC)


def _is_junction(path):
    if sys.platform != "win32" or not os.path.exists(path):
        return False
    import ctypes
    attr = ctypes.windll.kernel32.GetFileAttributesW(path)
    return attr != -1 and bool(attr & 0x400)  # FILE_ATTRIBUTE_REPARSE_POINT


def remove_existing():
    if os.path.islink(LINK_PATH) or _is_junction(LINK_PATH):
        if sys.platform == "win32":
            subprocess.run(["cmd", "/c", "rmdir", LINK_PATH], check=True)
        else:
            os.unlink(LINK_PATH)
    elif os.path.isdir(LINK_PATH):
        shutil.rmtree(LINK_PATH)
    elif os.path.exists(LINK_PATH):
        os.unlink(LINK_PATH)


def link():
    os.makedirs(TRA_SKILLS, exist_ok=True)
    if sys.platform == "win32":
        try:
            os.symlink(SKILL_SRC, LINK_PATH, target_is_directory=True)
            return "symlink"
        except OSError:
            # 没 symlink 权限时降级 junction（仅 Windows 本机可用）
            subprocess.run(["cmd", "/c", "mklink", "/J", LINK_PATH, SKILL_SRC],
                           check=True, capture_output=True)
            return "junction"
    else:
        os.symlink(SKILL_SRC, LINK_PATH)
        return "symlink"


def main():
    print(f"[源] {SKILL_SRC}")
    print(f"[目标] {LINK_PATH}")
    if already_ok():
        print("[跳过] 链接已存在且指向正确")
        return
    if os.path.lexists(LINK_PATH):
        remove_existing()
    kind = link()
    print(f"[完成] {kind}")
    if kind == "junction":
        print("[提示] junction 仅 Windows 本机可用，跨平台请开开发者模式后用 symlink")


if __name__ == "__main__":
    main()
