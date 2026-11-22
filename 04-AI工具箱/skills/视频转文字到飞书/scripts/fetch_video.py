"""
下载视频。抖音走 Playwright 抓直链，其他平台走 yt-dlp。

用法:
    python fetch_video.py <分享链接或视频URL> [输出目录]
"""
import os
import re
import sys
import glob
import time
import asyncio

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

# 短视频平台，需要 Playwright 抓直链
DOUYIN_HOSTS = ("douyin.com", "iesdouyin.com")


def is_douyin(url):
    return any(h in url for h in DOUYIN_HOSTS)


def resolve_short_link(url):
    """短链跳转后取出视频ID"""
    import requests
    r = requests.get(url, headers={"User-Agent": UA}, timeout=15, allow_redirects=True)
    final = r.url
    m = re.search(r'/(?:video|note)/(\d+)', final)
    if m:
        return m.group(1), final
    m = re.search(r'(\d{15,})', final)
    if m:
        return m.group(1), final
    return None, final


async def grab_douyin_url(video_id, timeout_ms=45000):
    """用浏览器打开抖音视频页，抓取真实视频流直链"""
    from playwright.async_api import async_playwright

    candidates = []

    async with async_playwright() as p:
        # 优先系统 Edge，避免下载 chromium
        try:
            browser = await p.chromium.launch(
                channel="msedge", headless=True,
                args=["--mute-audio", "--disable-blink-features=AutomationControlled"],
            )
        except Exception:
            browser = await p.chromium.launch(
                headless=True,
                args=["--mute-audio", "--disable-blink-features=AutomationControlled"],
            )

        context = await browser.new_context(
            user_agent=UA, viewport={"width": 1280, "height": 800}, locale="zh-CN")
        page = await context.new_page()

        def on_response(resp):
            u = resp.url
            if ".mp4" in u or "/video/tos/" in u:
                candidates.append(u)

        page.on("response", on_response)

        target = f"https://www.douyin.com/video/{video_id}"
        print(f"  打开: {target}")
        try:
            await page.goto(target, wait_until="domcontentloaded", timeout=timeout_ms)
        except Exception as e:
            print(f"  加载警告: {str(e)[:100]}")

        try:
            await page.wait_for_selector("video", timeout=20000)
            print("  video 元素已出现")
        except Exception:
            print("  未等到 video 元素，继续尝试")

        try:
            await page.evaluate(
                "() => { const v = document.querySelector('video');"
                " if (v) { v.muted = true; v.play().catch(()=>{}); } }")
        except Exception:
            pass

        await page.wait_for_timeout(6000)

        dom_srcs = await page.evaluate("""() => {
            const out = [];
            document.querySelectorAll('video').forEach(v => {
                if (v.src) out.push(v.src);
                if (v.currentSrc) out.push(v.currentSrc);
                v.querySelectorAll('source').forEach(s => { if (s.src) out.push(s.src); });
            });
            return out;
        }""")
        for s in dom_srcs:
            if s:
                candidates.append(s)

        try:
            title = await page.title()
        except Exception:
            title = ""

        await context.close()
        await browser.close()

    # 排序：真实视频流优先，douyinstatic 占位视频垫底
    ranked, seen = [], set()
    for c in candidates:
        if not c or c.startswith("blob:") or c in seen:
            continue
        seen.add(c)
        if "douyinvod.com" in c or "/video/tos/" in c:
            ranked.append((0, c))
        elif "douyin-pc-web" in c or "douyinstatic.com" in c:
            ranked.append((2, c))
        else:
            ranked.append((1, c))
    ranked.sort(key=lambda x: x[0])

    return [c for _, c in ranked], title


def download_stream(url, out_dir, referer="https://www.douyin.com/", name="video.mp4"):
    """下载单个直链，返回 (路径, 大小MB)"""
    import requests
    out = os.path.join(out_dir, name)
    r = requests.get(url, headers={"User-Agent": UA, "Referer": referer},
                     stream=True, timeout=60)
    r.raise_for_status()
    total = int(r.headers.get("content-length", 0))
    if total:
        print(f"  大小: {total/1024/1024:.1f} MB")
    with open(out, "wb") as f:
        for chunk in r.iter_content(65536):
            if chunk:
                f.write(chunk)
    size = os.path.getsize(out) / 1024 / 1024
    print(f"  已保存: {out} ({size:.1f} MB)")
    return out, size


def has_audio(path):
    """检查文件是否含音频轨。抖音返回的是分离流，视频流往往没有声音。"""
    try:
        import av
        container = av.open(path)
        ok = any(s.type == "audio" for s in container.streams)
        container.close()
        return ok
    except Exception:
        return False


def fetch_douyin(url, out_dir):
    print("[1/3] 解析视频ID...")
    m = re.search(r'/(?:video|note)/(\d+)', url)
    if m:
        video_id = m.group(1)
    else:
        video_id, final = resolve_short_link(url)
        print(f"  重定向: {final[:110]}")
    if not video_id:
        print("[失败] 无法解析视频ID")
        return None
    print(f"  ID: {video_id}")

    print("[2/3] 浏览器抓取视频直链...")
    urls, title = asyncio.run(grab_douyin_url(video_id))
    print(f"  标题: {title[:80]}")
    print(f"  候选链接 {len(urls)} 个")
    for u in urls[:5]:
        print(f"    {u[:110]}")

    if not urls:
        print("[失败] 未抓到直链")
        return None

    print("[3/3] 下载并探测音轨...")
    picked = None
    downloaded = []

    for i, cand in enumerate(urls[:4]):
        name = f"cand{i}.mp4"
        try:
            path, size = download_stream(cand, out_dir, name=name)
        except Exception as e:
            print(f"  第{i+1}个失败: {str(e)[:100]}")
            continue

        if size < 0.5:
            print(f"  占位视频({size:.2f}MB)，跳过")
            os.remove(path)
            continue

        audio = has_audio(path)
        downloaded.append(path)
        print(f"  轨道探测: {'含音频' if audio else '纯视频/无音轨'}")

        if audio:
            picked = path
            break

    if not picked and downloaded:
        # 没有音频轨的，退而取最大的那个（交给后续合并或直接失败）
        picked = max(downloaded, key=os.path.getsize)
        print(f"[警告] 未找到含音频的流，保留最大文件 {os.path.basename(picked)}")

    if not picked:
        print("[失败] 所有候选链接不可用")
        return None

    final = os.path.join(out_dir, "video.mp4")
    if os.path.abspath(picked) != os.path.abspath(final):
        if os.path.exists(final):
            os.remove(final)
        os.replace(picked, final)

    for p in downloaded:
        if os.path.exists(p) and os.path.abspath(p) != os.path.abspath(final):
            os.remove(p)

    print(f"[完成] {final} ({os.path.getsize(final)/1024/1024:.1f} MB)")
    return final


def fetch_ytdlp(url, out_dir):
    """B站/小红书/YouTube 等走 yt-dlp"""
    import yt_dlp

    outtmpl = os.path.join(out_dir, "video.%(ext)s")
    opts = {
        "outtmpl": outtmpl,
        "format": "best",
        "quiet": True,
        "no_warnings": True,
        "socket_timeout": 30,
        "http_headers": {"User-Agent": UA},
    }

    print("[1/2] yt-dlp 下载...")
    try:
        with yt_dlp.YoutubeDL(opts) as ydl:
            info = ydl.extract_info(url, download=True)
    except Exception as e:
        print(f"[失败] {str(e).split(chr(10))[0][:250]}")
        return None

    files = [f for f in glob.glob(os.path.join(out_dir, "video.*"))
             if not f.endswith((".part", ".ytdl"))]
    if not files:
        print("[失败] 未生成文件")
        return None

    path = max(files, key=os.path.getmtime)
    print(f"  标题: {info.get('title', 'N/A')}")
    print(f"  时长: {info.get('duration', 0)} 秒")
    print(f"  作者: {info.get('uploader', 'N/A')}")
    print(f"[完成] {path} ({os.path.getsize(path)/1024/1024:.1f} MB)")
    return path


def main():
    if len(sys.argv) < 2:
        print("用法: python fetch_video.py <url> [out_dir]")
        sys.exit(1)

    url = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()
    os.makedirs(out_dir, exist_ok=True)

    path = fetch_douyin(url, out_dir) if is_douyin(url) else fetch_ytdlp(url, out_dir)
    sys.exit(0 if path else 1)


if __name__ == "__main__":
    main()
