---
name: video-transcript
description: Transcribe Douyin or Bilibili share links into text and write the result into a Feishu doc. Use when the user pastes a short video link and asks to transcribe it, or says 视频转文字, 转录, or 写成文档.
---

# Video Transcript

把视频链接转成逐字稿，并落到飞书文档。

## 流程

```
分享链接 → 解析视频ID → 下载音频流 → 本地转录 → 写飞书文档
```

| 步骤 | 脚本 | 产出 |
|---|---|---|
| 1. 下载 | `scripts/fetch_video.py <链接> [目录]` | `video.mp4`（含音轨） |
| 2. 转录 | `scripts/local_asr.py <视频> [目录]` | `transcript.txt` |
| 3. 成文 | 读 `transcript.txt`，用 `lark-cli docs +create` 写入 | 飞书文档 |

**默认走本地转录，零成本无限量。** 妙记只在本地不可用时兜底（见下）。

## 环境准备

```powershell
python -c "import yt_dlp, playwright, requests, av, faster_whisper"
# 缺哪个装哪个（国内用清华镜像）
python -m pip install -U yt-dlp playwright requests av faster_whisper -i https://pypi.tuna.tsinghua.edu.cn/simple
# GPU 加速所需的 CUDA 运行库（不用装完整 CUDA Toolkit）
python -m pip install nvidia-cublas-cu12 nvidia-cudnn-cu12 -i https://pypi.tuna.tsinghua.edu.cn/simple
```

`playwright` 库即可，**不需要** `playwright install chromium` —— 脚本复用系统 Edge。

### 模型准备（首次）

```powershell
python scripts/dl_model.py large-v3
```

模型落在 `D:\models\fw-<size>`，可用环境变量 `FW_MODEL_DIR` 改根目录。

- **一定走 ModelScope**，实测 8.5 MB/s；HuggingFace 直连只有 0.16 MB/s（3GB 要下 5 小时）。
- large-v3 约 3GB / 6 分钟，中文效果最好；medium 约 1.5GB 可作折中。
- 下载一次长期使用，之后转录完全免费。

## 关键技术判断

这些都是实测踩过的坑：

1. **yt-dlp 的抖音提取器不可用**。源码 `extractor/tiktok.py` 的抖音分支只有 `TODO: Run verification challenge code`，签名挑战未实现，给任何 cookie 都报 `Fresh cookies (not necessarily logged in) are needed`。不要在这条路上折腾。

2. **抖音必须用真实浏览器抓直链**。`__ac_signature` 由前端 JS 生成，纯 `requests` 拿不到；服务端渲染的 `_ROUTER_DATA` 里也没有视频字段（客户端异步加载）。用 Playwright 打开 `https://www.douyin.com/video/<id>`，监听 `response` 抓 `.mp4` / `/video/tos/`。

3. **Chrome 运行时会锁 cookie 数据库**。yt-dlp 报 `Could not copy Chrome cookie database` 即此。用 Edge（`channel="msedge"`）或先关浏览器。

4. **抖音返回的是分离的音视频流**，最容易踩的坑。同一视频的多个候选链，有的是纯视频（85MB / 1080p / 无音轨），有的是纯音频（22MB / aac / 无画面）。**转录只需音频**，下载后必须用 `av.open` 探测音轨取含 audio 的那份。否则会把纯视频喂给 ASR，得到一堆静音或无输出。

5. **还要过滤占位视频**。候选常混入 `douyinstatic.com/.../uuu_265.mp4`（约 0.2MB 预加载占位）。小于 0.5MB 的一律丢弃。

6. **CUDA DLL 要手动注入**。pip 装的 `nvidia-cublas-cu12` / `nvidia-cudnn-cu12` 不会进 PATH，直接跑会报 `Library cublas64_12.dll is not found`。需在 `import faster_whisper` 前把这些目录加入 `PATH` 和 `os.add_dll_directory`。

7. **HF 缓存里的模型可能是坏的**。`snapshots/*/model.bin` 常是 0 字节的失效链接，真实文件在 `blobs/` 下。别看到目录存在就以为能用，要校验 `model.bin` 实际大小。

8. **large-v3 必须有 `preprocessor_config.json`**。faster-whisper 从这个文件读 `feature_size`，**不是** config.json。缺了会默认按 80 维处理，然后报 `expected input with shape (1, 128, 3000), but got (1, 80, 3000)`。ModelScope 的镜像仓库没有这个文件，需自行写入 `{"feature_size": 128, ...}`。

9. **ModelScope 的 large-v3 仓库缺 `vocabulary.txt`**（404），缺了报 `Cannot load the vocabulary from the model directory`。Whisper 多语言模型共用同一词表，从任意本地 whisper 模型复制即可，或从 HF 单独拉这个 460KB 文件。

## 妙记兜底（本地不可用时）

每人每月 **300 分钟**语音转文字，超出后当月不再生成逐字稿，付费 ¥69/月。批量转录会很快耗尽，所以定位是兜底而非主力。

```powershell
python scripts/feishu_asr.py <视频> [目录]
```

注意两点：
- `--wait-ready` 返回时逐字稿可能还没生成，只拿到时长和空 Keywords，要按正文长度轮询重试
- `lark-cli` 的 JSON 可能输出到 stderr，解析时两边都要试

## 性能与质量参考

RTX 2080，15.9 分钟音频实测：

| 模型 | 耗时 | 速度 | 质量 |
|---|---|---|---|
| large-v3 (int8_float16) | 225s | 4.25x 实时 | 好，术语基本正确 |
| small | 122s | 7.78x 实时 | 快但错字多，不推荐 |

small 会把"望闻问切"识别成"忘文问妾"、"批量化"成"劈炼化"、"论坛"成"令台"。large-v3 全部修正。**默认用 large-v3**，快不是理由。

妙记约 60s 出稿，质量与 large-v3 相当，但受免费额度限制。

## 术语修正

`scripts/glossary.json` 存放错字映射（`{"错写":"正写"}`），转录后自动替换。领域的专有名词应逐步积累到这个文件。

## 输出约定

- 逐字稿保留时间轴分段，不要合并成一段
- 飞书文档结构：视频信息（平台/作者/链接/时长）→ 内容主线提炼 → 转录全文
- 提炼只写原作者的真实论点，不掺入自己的判断

## 边界

- 只处理用户明确给出的链接，不主动批量爬取
- 不绕过登录墙：需要登录的内容先请用户自己登录
