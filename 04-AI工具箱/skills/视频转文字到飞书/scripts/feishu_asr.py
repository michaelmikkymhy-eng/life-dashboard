"""
把音视频文件转成逐字稿：上传飞书云空间 -> 生成妙记 -> 取逐字稿。

用法:
    python feishu_asr.py <音视频文件> [输出目录]
"""
import os
import sys
import json
import glob
import time
import subprocess

def find_lark_cli():
    """按优先级找 lark-cli：环境变量 LARK_CLI > PATH > 常见插件目录"""
    env = os.environ.get('LARK_CLI')
    if env and os.path.exists(env):
        return env

    # PATH 里有就直接用
    from shutil import which
    p = which('lark-cli') or which('lark-cli.exe')
    if p:
        return p

    # 兜底：在用户目录下几个已知插件根里 glob，取版本号最大的
    home = os.path.expanduser('~')
    roots = [
        os.path.join(home, '.trae-cn', 'plugins', 'trae-remote-official', 'lark'),
        os.path.join(home, '.trae', 'plugins', 'trae-remote-official', 'lark'),
    ]
    for base in roots:
        if not os.path.isdir(base):
            continue
        found = glob.glob(os.path.join(base, '**', 'lark-cli.exe'), recursive=True)
        if found:
            return sorted(found)[-1]
    return None


def run(args, cwd=None):
    return subprocess.run(args, capture_output=True, text=True,
                          encoding='utf-8', errors='replace', cwd=cwd)


def parse_ok(stdout, stderr):
    """lark-cli 的 JSON 可能落在 stdout 或 stderr，两边都试"""
    for out in (stdout, stderr):
        if not out:
            continue
        out = out.strip()
        if not out.startswith('{'):
            continue
        try:
            data = json.loads(out)
        except Exception:
            continue
        if data.get('ok'):
            return data
        return {'_error': json.dumps(data, ensure_ascii=False)[:400]}
    return None


def main():
    if len(sys.argv) < 2:
        print("用法: python feishu_asr.py <file> [out_dir]")
        sys.exit(1)

    media = os.path.abspath(sys.argv[1])
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(media)

    if not os.path.exists(media):
        print(f"[失败] 文件不存在: {media}")
        sys.exit(1)

    lark = find_lark_cli()
    if not lark:
        print("[失败] 找不到 lark-cli，先确认飞书插件已安装")
        sys.exit(1)
    print(f"[环境] {lark}")

    size_mb = os.path.getsize(media) / 1024 / 1024

    print(f"[1/3] 上传到飞书云空间 ({size_mb:.1f}MB)...")
    r = run([lark, 'drive', '+upload', '--file', media,
             '--name', os.path.basename(media)], cwd=out_dir)
    data = parse_ok(r.stdout, r.stderr)
    if not data or data.get('_error'):
        msg = data.get('_error') if data else (r.stdout[:300] + r.stderr[:300])
        print(f"[失败] 上传: {msg}")
        sys.exit(1)
    file_token = data['data'].get('file_token')
    print(f"  file_token: {file_token}")

    print("[2/3] 生成妙记...")
    r = run([lark, 'minutes', '+upload', '--file-token', file_token], cwd=out_dir)
    data = parse_ok(r.stdout, r.stderr)
    if not data or data.get('_error'):
        msg = data.get('_error') if data else (r.stdout[:300] + r.stderr[:300])
        print(f"[失败] 妙记: {msg}")
        sys.exit(1)
    minute_token = data['data'].get('minute_token')
    minute_url = data['data'].get('minute_url', '')
    print(f"  minute_token: {minute_token}")
    print(f"  妙记链接: {minute_url}")

    print("[3/3] 等待妙记生成并取逐字稿...")
    r = run([lark, 'minutes', '+detail', '--minute-tokens', minute_token,
             '--wait-ready', '--transcript', '--output-dir', out_dir], cwd=out_dir)
    data = parse_ok(r.stdout, r.stderr)
    if not data or data.get('_error'):
        msg = data.get('_error') if data else (r.stdout[:300] + r.stderr[:300])
        print(f"[失败] 详情: {msg}")
        sys.exit(1)

    minutes = data['data'].get('minutes', [])
    if not minutes:
        print("[失败] 返回中没有 minutes")
        sys.exit(1)

    m = minutes[0]
    print(f"  标题: {m.get('title', '')}")

    def read_transcript(minute_obj):
        """从 detail 结果里取逐字稿正文"""
        art = minute_obj.get('artifacts', {}) or {}
        tfile = art.get('transcript_file', '')
        if tfile and os.path.exists(tfile):
            with open(tfile, 'r', encoding='utf-8') as f:
                return f.read()
        return art.get('transcript', '') or ''

    def body_len(text):
        """去掉表头后统计正文字数，用来判断是否真的转完了"""
        import re
        body = re.sub(r'^\d{4}-\d{2}-\d{2}[^\n]*\n', '', text)
        body = re.sub(r'^Keywords:.*?(\n\n|\Z)', '', body, flags=re.S)
        return len(re.sub(r'\s', '', body))

    text = read_transcript(m)

    # 妙记可能先返回空稿，正文为空时轮询等待
    max_wait, interval = 600, 20
    waited = 0
    while body_len(text) < 50 and waited < max_wait:
        print(f"  正文仍为空(已等{waited}s)，{interval}s 后重试...")
        time.sleep(interval)
        waited += interval
        r = run([lark, 'minutes', '+detail', '--minute-tokens', minute_token,
                 '--transcript', '--output-dir', out_dir], cwd=out_dir)
        d2 = parse_ok(r.stdout, r.stderr)
        if d2 and not d2.get('_error'):
            ms = d2['data'].get('minutes', [])
            if ms:
                text = read_transcript(ms[0])

    if body_len(text) < 50:
        print(f"[失败] 等待 {max_wait}s 后逐字稿仍为空，妙记可能还在处理")
        print(f"  可稍后手动查看: {minute_url}")
        sys.exit(1)

    dest = os.path.join(out_dir, 'transcript.txt')
    with open(dest, 'w', encoding='utf-8') as f:
        f.write(text)

    print(f"[完成] 逐字稿 {len(text)} 字 -> {dest}")
    print(f"[妙记] {minute_url}")


if __name__ == '__main__':
    main()
