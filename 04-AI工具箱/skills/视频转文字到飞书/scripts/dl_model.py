"""
从 ModelScope 下载 faster-whisper 模型（国内 9MB/s 级别，远快于 HuggingFace）

用法:
    python dl_model.py [large-v3|medium]
"""
import os
import sys
import json
import time
import shutil
import requests

REPOS = {
    "large-v3": "pengzhendong/faster-whisper-large-v3",
    "medium": "pengzhendong/faster-whisper-medium",
}
# CTranslate2 模型运行所需文件；vocabulary.txt 部分仓库缺失，需另行补齐
FILES = ["config.json", "model.bin", "tokenizer.json"]
OPTIONAL = ["vocabulary.txt"]
UA = {"User-Agent": "Mozilla/5.0"}
MODEL_ROOT = os.environ.get("FW_MODEL_DIR", r"D:\models")

# large-v3 用 128 维 mel 特征，其余 80。faster-whisper 从 preprocessor_config.json 读取，
# 缺这个文件会默认按 80 处理，导致 "expected (1,128,3000) but got (1,80,3000)"
MEL_BINS = {"large-v3": 128, "medium": 80}


def download(url, out, label):
    t0 = time.time()
    r = requests.get(url, headers=UA, stream=True, timeout=60)
    r.raise_for_status()
    total = int(r.headers.get("content-length", 0))
    tmp = out + ".part"
    got = last = 0
    with open(tmp, "wb") as f:
        for chunk in r.iter_content(1048576):
            if not chunk:
                continue
            f.write(chunk)
            got += len(chunk)
            if total and got - last >= 200 * 1024 * 1024:
                last = got
                sp = got / 1024 / 1024 / max(time.time() - t0, .01)
                print(f"  {got/total*100:.0f}% "
                      f"({got/1024/1024:.0f}/{total/1024/1024:.0f}MB) {sp:.1f}MB/s")
    os.replace(tmp, out)
    print(f"  完成 {os.path.getsize(out)/1024/1024:.1f}MB ({time.time()-t0:.0f}s)")


def ensure_vocabulary(dest, size):
    """vocabulary.txt 部分镜像仓库没有。Whisper 多语言模型共用同一词表，
    可从其他本地模型复制，或从 HF 拉取（文件仅 460KB）"""
    out = os.path.join(dest, "vocabulary.txt")
    if os.path.exists(out) and os.path.getsize(out) > 1000:
        return True

    for other in ("fw-small", "fw-medium", "fw-large-v3"):
        src = os.path.join(MODEL_ROOT, other, "vocabulary.txt")
        if os.path.abspath(src) == os.path.abspath(out):
            continue
        if os.path.exists(src) and os.path.getsize(src) > 1000:
            shutil.copy2(src, out)
            print(f"  vocabulary.txt 从 {other} 补齐 "
                  f"({os.path.getsize(out)/1024:.0f}KB)")
            return True

    url = ("https://huggingface.co/Systran/faster-whisper-large-v3"
           "/resolve/main/vocabulary.txt")
    try:
        r = requests.get(url, headers=UA, timeout=60)
        if r.status_code == 200:
            with open(out, "wb") as f:
                f.write(r.content)
            print(f"  vocabulary.txt 从 HuggingFace 补齐 "
                  f"({os.path.getsize(out)/1024:.0f}KB)")
            return True
    except Exception as e:
        print(f"  补齐 vocabulary.txt 失败: {str(e)[:80]}")
    return False


def write_preprocessor(dest, size):
    """写入 mel 特征配置，缺了会导致维度不匹配"""
    cfg = {
        "feature_size": MEL_BINS.get(size, 80),
        "sampling_rate": 16000,
        "hop_length": 160,
        "chunk_length": 30,
        "n_fft": 400,
        "padding_value": 0.0,
        "nb_max_frames": 3000,
    }
    p = os.path.join(dest, "preprocessor_config.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    print(f"  preprocessor_config.json 已写入 (feature_size={cfg['feature_size']})")


def main():
    size = sys.argv[1] if len(sys.argv) > 1 else "large-v3"
    repo = REPOS.get(size)
    if not repo:
        sys.exit(f"未知模型 {size}，可选: {list(REPOS)}")

    dest = os.path.join(MODEL_ROOT, f"fw-{size}")
    os.makedirs(dest, exist_ok=True)
    print(f"[目标] {dest}")
    print(f"[仓库] {repo}")

    for fn in FILES:
        out = os.path.join(dest, fn)
        if os.path.exists(out) and os.path.getsize(out) > 0:
            print(f"[跳过] {fn} ({os.path.getsize(out)/1024/1024:.1f}MB)")
            continue
        print(f"[下载] {fn} ...")
        try:
            download(f"https://www.modelscope.cn/models/{repo}/resolve/master/{fn}", out, fn)
        except Exception as e:
            print(f"  失败: {str(e)[:120]}")
            sys.exit(1)

    print("[补齐] vocabulary.txt ...")
    ensure_vocabulary(dest, size)

    print("[配置] preprocessor_config.json ...")
    write_preprocessor(dest, size)

    print("\n[完成] 文件清单")
    for fn in sorted(os.listdir(dest)):
        p = os.path.join(dest, fn)
        if os.path.isfile(p):
            print(f"  {fn} {os.path.getsize(p)/1024/1024:.2f}MB")


if __name__ == "__main__":
    main()
