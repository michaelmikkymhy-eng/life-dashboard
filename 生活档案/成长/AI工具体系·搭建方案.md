# AI 工具体系搭建方案（Hermes + Grok 中转 + 本地知识库）

> 本文件是施工图，不是学习资料。每一步都有「做完的标志」，没达到就不往下走。
> 依据：Hermes Agent 官方文档（hermes-agent.nousresearch.com，2026-11）+ 抖音四步法

---

## 0. 先明确一件事（重要）

**Hermes Agent 是跑在你终端里的自主智能体，不是网页服务。**它的能力来自"能读写你的文件 + 能调工具"，所以：

- ✅ 它天然适合当"本地知识库的操作系统"
- ❌ 它不能替你在网页版 Grok / ChatGPT 之间点按钮——"中转"发生在 API 层，不在网页层

所以你说的"以 Grok 为中转"，实际落地形态是：**Hermes 是大脑和调度器，Grok 是它默认调用的大模型**。其他 AI（豆包、DeepSeek、Claude）作为**辅助模型槽位**，被 Hermes 按任务类型分流。这不是"让几个 AI 互相聊天"，而是"一个调度器 + 多个模型按槽位分工"——后者才是能稳定跑的。

---

## 1. 目标（第 1 步：定目标）

**11 月 30 日前，达成：**

| 验收项 | 通过标准 |
|---|---|
| ① Hermes 能跑 | 终端里 `hermes` 能进对话，回答里能显示当前模型 |
| ② Grok 已接入 | `/model` 显示 Grok，且一次真实提问有正常回复 |
| ③ 能读本地库 | 让它"总结 生活档案/日记/2026-11-22.md"，能准确复述内容 |
| ④ 能写本地库 | 让它把一段内容写进 `知识库/收件箱.md`，文件真的出现 |
| ⑤ 辅助模型分工 | 至少 1 个辅助槽位用国产模型（豆包/DeepSeek），且不影响主对话 |

**五条全过 = 体系成立。**任何一条没过，先修这一条，别加功能。

---

## 2. 架构（一张图看懂）

```
你（终端 / Telegram）
        │
        ▼
   Hermes Agent  ←── 技能库、cron、记忆（~/.hermes/）
        │
        ├── 主模型槽位：Grok 4（统筹、长任务、工具调用）
        │
        ├── 辅助槽位 A：豆包/火山（便宜、中文日常、上下文压缩）
        ├── 辅助槽位 B：DeepSeek（代码、结构化）
        └── 辅助槽位 C：视觉模型（看图、截图理解）
        │
        ▼
   本地知识库（纯文本，git 管理）
   /生活档案   ← 已有，日记+台账
   /知识库     ← 新建，AI 学习沉淀
```

**关键设计原则**：知识库永远是纯 Markdown + git。这样 Hermes 挂了、换 Claude Code、换 Cursor，积累都还在——这就是视频第 4 步"换工具也能继续"。

---

## 3. 施工步骤

### Step 1｜装 Hermes（30 分钟）

macOS 上二选一：

**A. 只装 CLI（推荐，先跑通再说）**
```bash
curl -fsSL https://hermes-agent.nousresearch.com/install.sh | bash
source ~/.zshrc
hermes --version
```

**B. 装桌面版（含 CLI）**
去 https://hermes-agent.nousresearch.com/ 下载 Hermes Desktop 安装器。

⚠️ 官方明确说：**不要用 `pip install hermes-agent`**，必须用上面的受管安装脚本。

**做完的标志**：`hermes --version` 有版本号输出。

### Step 2｜接 Grok（20 分钟）

你有两条路，取决于账号类型：

**路线 A：有 SuperGrok / Premium+ 订阅（省钱，推荐）**
```bash
hermes model
# 选 "xAI Grok OAuth (SuperGrok / Premium+)"
# 浏览器会弹出登录，授权完回终端
```

**路线 B：只有 API key**
```bash
hermes config set XAI_API_KEY xai-你的key
hermes model        # 选 xai 提供商 → 选 grok-4-1-fast-reasoning 之类
```

**注意**：Hermes 要求模型**上下文 ≥ 64K token**。Grok 4 系列满足，老的 grok-2 不满足，会被直接拒绝启动。

**做完的标志**：`hermes` 进入对话，问一句"你现在用哪个模型"，它答得出 Grok。

### Step 3｜接辅助模型（20 分钟）

你已经有火山 API 套餐（¥200），可以直接当辅助槽位：

```bash
hermes config set VOLC_API_KEY 你的火山key
hermes model    # 辅助模型 → Custom endpoint 或对应 provider
```

火山方舟的 endpoint 是 OpenAI 兼容的，base_url 形如 `https://ark.cn-beijing.volces.com/api/v3`，模型名要填你开通的 **endpoint id**（形如 `ep-2024xxxx`），不是模型中文名。**这是最容易卡住的一步**，卡了就把报错发我。

辅助槽位建议只配 2 个，别贪多：**豆包（中文/便宜）+ DeepSeek（代码）**。

**做完的标志**：主动触发一次上下文压缩（长对话），观察它是否用了辅助模型。

### Step 4｜连本地知识库（30 分钟）

知识库目录（`/workspace/生活档案` 已存在，直接复用）：

```bash
cd /workspace
mkdir -p 知识库
cd 知识库
cat > 目标.md <<'EOF'
# 目标
- 11/30 前：Hermes + Grok 跑通，知识库可读写
EOF
cat > 错题本.md <<'EOF'
# 错题本
| 日期 | 卡在哪 | 报错原文 | 怎么解决的 |
|---|---|---|---|
EOF
cat > 进度.md <<'EOF'
# 进度
| 周 | 本周产物 | 能不能跑 |
|---|---|---|
EOF
```

然后让 Hermes 读一次，验证它真的能访问：

> 「读一下 生活档案/日记/2026-11-22.md，用 3 条总结我今天做了什么」

**做完的标志**：它复述的内容和日记一致（不是编的）。

### Step 5｜建第一个技能（30 分钟）

Hermes 的核心是"从经验里生成技能"。第一个技能就做你天天要做的事：

> 「我每天要记录饮食和作息，你把这件事做成一个技能：我口述一段流水，你解析成结构化的时间轴、饮食、称重记录，写进 生活档案/日记/当天日期.md，格式参考已有文件。」

**做完的标志**：说一段测试流水，它能生成格式正确的文件。

### Step 6｜留在 Git 里（10 分钟）

```bash
cd /workspace
git add 知识库 && git commit -m "建立知识库骨架"
```

每周日 `git log --oneline -7` 看一眼，这就是你的"学习进度条"。

---

## 4. 安装卡点自查表（你说之前没装上，先从这查）

| 症状 | 大概率原因 | 处理 |
|---|---|---|
| `curl: command not found` | 极简环境 | macOS 自带 curl，说明 PATH 异常，`echo $PATH` 发我 |
| 安装脚本跑到一半失败 | 网络到 hermes-agent.nousresearch.com 不通 | 挂代理重跑，或改用 Desktop 安装器 |
| `hermes: command not found`（装完） | 没重载 shell | `source ~/.zshrc`，或重开终端；仍不行看 `~/.hermes/bin` 是否在 PATH |
| 能进对话但一提问就报错 | 没配 provider 或 key 无效 | `hermes config` 检查，`hermes model` 重选 |
| 报 "context window too small" | 模型上下文 < 64K | 换 Grok 4 / Claude / GPT 系列 |
| 卡在登录授权 | 浏览器回调端口被占 | 换终端重试，或改用 API key 路线 |
| 火山接口 404/401 | 模型名填了中文名而非 endpoint id | 去火山控制台复制 `ep-` 开头的 id |

---

## 5. 明确不做什么（防止跑偏）

1. **不同时装三套 agent**（Hermes + Claude Code + Cursor），先跑通一个
2. **不先接 Telegram/Discord gateway**——CLI 对话正常之后再谈远程触达
3. **不追求"全自动"**——先做到"你说一句，它记一条"，再谈自动化
4. **不把 key 写进任何会被 git 提交的文件**（`~/.hermes/.env` 不在仓库里，安全）

---

*版本依据：Hermes Agent 官方文档 2026-11 版 ｜ 建立日期：2026-11-22*
