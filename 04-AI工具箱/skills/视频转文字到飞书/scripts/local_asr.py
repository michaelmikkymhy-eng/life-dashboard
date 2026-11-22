"""
本地 faster-whisper 转录（免费，GPU 加速）

用法:
    python local_asr.py <音视频文件> [输出目录]

模型目录约定: D:\\models\\fw-<size>，可用环境变量 FW_MODEL_DIR 覆盖。
优先 large-v3，依次回退 medium / small。
"""
import os
import sys
import json
import time

MODEL_ROOT = os.environ.get("FW_MODEL_DIR", r"D:\models")
PREFER = ["large-v3", "medium", "small"]
GLOSSARY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "glossary.json")


def inject_cuda_dlls():
    """pip 装的 nvidia CUDA 运行库不在 PATH 里，手动注入"""
    site = os.path.join(os.path.dirname(os.__file__), "site-packages", "nvidia")
    added = []
    for sub in ("cublas", "cudnn", "cuda_runtime", "cuda_nvrtc"):
        p = os.path.join(site, sub, "bin")
        if os.path.isdir(p):
            os.environ["PATH"] = p + os.pathsep + os.environ.get("PATH", "")
            try:
                os.add_dll_directory(p)
            except Exception:
                pass
            added.append(sub)
    return added


def find_model():
    """返回第一个存在的本地模型目录"""
    for size in PREFER:
        d = os.path.join(MODEL_ROOT, f"fw-{size}")
        if os.path.isdir(d) and os.path.exists(os.path.join(d, "model.bin")):
            mb = os.path.getsize(os.path.join(d, "model.bin")) / 1024 / 1024
            if mb > 10:
                return d, size, mb
    return None, None, 0


def load_glossary():
    if not os.path.exists(GLOSSARY):
        return {}
    try:
        with open(GLOSSARY, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def main():
    if len(sys.argv) < 2:
        print("用法: python local_asr.py <file> [out_dir]")
        sys.exit(1)

    media = os.path.abspath(sys.argv[1])
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(media)
    if not os.path.exists(media):
        sys.exit(f"[失败] 文件不存在: {media}")

    cuda = inject_cuda_dlls()
    print(f"[CUDA] 注入: {cuda or '无（将走 CPU）'}")

    model_dir, size, mb = find_model()
    if not model_dir:
        print(f"[失败] 在 {MODEL_ROOT} 下没找到 fw-large-v3 / fw-medium / fw-small")
        print("  下载命令: python scripts/dl_model.py large-v3")
        sys.exit(1)
    print(f"[模型] {size} ({mb:.0f}MB) @ {model_dir}")

    from faster_whisper import WhisperModel

    model, used = None, None
    for device, ctype in [("cuda", "int8_float16"), ("cuda", "float16"),
                          ("cuda", "int8"), ("cpu", "int8")]:
        try:
            print(f"[加载] {device}/{ctype} ...")
            t0 = time.time()
            model = WhisperModel(model_dir, device=device, compute_type=ctype)
            print(f"  成功（{time.time()-t0:.1f}s）")
            used = f"{device}/{ctype}"
            break
        except Exception as e:
            print(f"  失败: {str(e).split(chr(10))[0][:140]}")
    if model is None:
        sys.exit("[失败] 所有设备都无法加载")

    print(f"[转录] {used}")
    t0 = time.time()
    segments, info = model.transcribe(
        media, language="zh", beam_size=5, vad_filter=True,
        vad_parameters={"min_silence_duration_ms": 2000,
                        "speech_pad_ms": 400, "threshold": 0.35},
        initial_prompt="以下是中文口播内容，可能涉及传统文化、AI工具、职场与个人成长。",
    )

    print(f"  语言 {info.language} ({info.language_probability:.1%})，"
          f"时长 {info.duration/60:.1f} 分钟")

    lines = []
    for seg in segments:
        t = seg.text.strip()
        if t:
            lines.append(f"[{seg.start:.1f}-{seg.end:.1f}] {t}")

    text = "\n".join(lines)

    # 术语后处理
    gl = load_glossary()
    if gl:
        n = 0
        for wrong, right in gl.items():
            c = text.count(wrong)
            if c:
                text = text.replace(wrong, right)
                n += c
        print(f"[术语] 修正 {n} 处（词典 {len(gl)} 条）")

    el = time.time() - t0
    print(f"[结果] {len(text)} 字，耗时 {el:.1f}s，"
          f"{info.duration/max(el,0.1):.2f}x 实时")

    dest = os.path.join(out_dir, "transcript.txt")
    with open(dest, "w", encoding="utf-8") as f:
        f.write(f"# 本地转录 ({size})\n# {time.strftime('%Y-%m-%d %H:%M:%S')}"
                f" | {info.duration/60:.1f}min\n\n{text}\n")
    print(f"[完成] {dest}")


if __name__ == "__main__":
    main()
