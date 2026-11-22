# AI 工具体系搭建方案（Windows + Hermes + 本地知识库）

> 本文件是施工图，不是学习资料。每一步都有「做完的标志」，没达到就不往下走。
> 依据：Hermes Agent 官方文档（hermes-agent.nousresearch.com，2026-11 版）+ 抖音四步法

---

## 0. 三个必须先纠正的认知

### ① Hermes 是终端里的智能体，不是"AI 之间的中转站"

它不是让几个 AI 互相聊天。真实结构是：**Hermes 做调度器 + 记忆 + 工具执行，底层挂模型**。所谓"以 Grok 为中转"，落地形态是"Grok 当 Hermes 的主模型"。其他 AI 挂在**辅助模型槽位**上，按任务类型分流（便宜任务走便宜模型）。

### ② Windows 能装，两条路都行

官方支持 Windows 原生安装，**不需要 WSL**：

```powershell
# 方式 A：只装 CLI（推荐先跑通）
iex (irm https://hermes-agent.nousresearch.com/install.ps1)

# 方式 B：装桌面版（含 CLI）
# 去 https://hermes-agent.nousresearch.com/ 下载 Windows 安装器
```

装完开新的 PowerShell 窗口（不是重载当前窗口），然后 `hermes --version`。

### ③ 免费版 Grok 接不进来，这是硬限制

Hermes 接 Grok 只有两条路：**SuperGrok/Premium+ 订阅走 OAuth**，或**xAI API key 按量付费**。免费账号只有网页版 grok.com，没有 API 访问权——**技术上无法接入**。

所以现在的现实选择是：

| 方案 | 主模型 | 花费 | 可行性 |
|---|---|---|---|
| **A（建议）** | 火山豆包（你已付 ¥200） | 已付 | ✅ 立刻能跑 |
| **B** | SuperGrok 订阅 + OAuth | 月费 | 需要先付费 |
| **C** | xAI API key | 按 token 计费 | 需要充值 |

**建议先走 A**：你手上的火山套餐是沉没成本，先用起来。Grok 等你确认要付费再当主力接进来——Hermes 支持随时 `hermes model` 切换，没有锁定。

---

## 1. 目标（第 1 步：定目标）

**11 月 30 日前达成：**

| 验收项 | 通过标准 |
|---|---|
| ① Hermes 能跑 | PowerShell 里 `hermes` 能进对话 |
| ② 主模型已接 | 对话时能显示当前模型，回答正常 |
| ③ 能读本地库 | 让它总结 `01-日常管理\日记\2026-11-22.md`，内容准确 |
| ④ 能写本地库 | 让它写进 `02-意识与创作\灵感\收件箱.md`，文件真的出现 |
| ⑤ 辅助模型分工 | 至少 1 个辅助槽位在跑 |

五条全过 = 体系成立。任何一条没过，先修这一条，别加功能。

---

## 2. 架构

```
你（Windows 终端 / 后续可接 Telegram）
        │
        ▼
   Hermes Agent（技能库、记忆、cron）→ %USERPROFILE%\.hermes\
        │
        ├── 主模型槽位：豆包（火山，已付费）→ 后期可切 Grok
        ├── 辅助槽位 A：DeepSeek（代码、结构化，极便宜）
        └── 辅助槽位 B：视觉/压缩（豆包同源即可）
        │
        ▼
   本地 Obsidian 镜像（Markdown + Excel + git）
   年轮\01-日常管理   ← 云端 AI 维护的日记与唯一总台账
   年轮\02-意识与创作 ← 本地可读写的灵感、随笔与精神成长
   年轮\04-AI工具箱   ← 可迁移工具、Skill 与 prompts
```

**核心原则**：年轮以 Markdown + git 为交换层，生活数据总台账始终只有一份 Excel。Hermes 挂了、换 Claude Code、换 Cursor，积累都还在——这就是视频第 4 步"换工具也能继续"。日常日记与 Excel 只允许云端 AI 通过 sync.py 双写；本地 Hermes 可以读取，也可以写灵感、随笔和项目文档，但不能直接修改日常管理数据。

---

## 3. 施工步骤

### Step 1｜Windows 装 Hermes

```powershell
iex (irm https://hermes-agent.nousresearch.com/install.ps1)
```

装完**关掉 PowerShell 重新开**，然后：
```powershell
hermes --version
```

⚠️ 官方明确禁止 `pip install hermes-agent`，必须用上面的脚本。

**做完的标志**：`hermes --version` 出版本号。

### Step 2｜把知识库拉到 Windows 本地

你这套档案在 GitHub 私有仓库 `life-dashboard` 里（342 次提交，我这边一直在同步）。Windows 上拉一份：

```powershell
cd $env:USERPROFILE
git clone https://github.com/michaelmikkymhy-eng/life-dashboard.git 年轮
cd 年轮
git config core.hooksPath .githooks
```

> 仓库是私有的，需要登录凭据。用 Git Credential Manager 弹窗登录 GitHub 即可，**不要把 token 贴进任何文件**。

**做完的标志**：`dir 年轮\01-日常管理\日记` 能看到 md 文件。

### Step 3｜接主模型（豆包 / 火山）

```powershell
hermes model
# 选 Custom endpoint（OpenAI 兼容）
# base_url: https://ark.cn-beijing.volces.com/api/v3
# api key : 火山控制台的 API Key
# model   : 你的接入点 ID，形如 ep-2024xxxx（不是中文模型名！）
```

**这是最容易卡的一步**：火山要求填 **endpoint id（`ep-` 开头）**，填中文名必 404。上下文要 ≥ 64K，选支持长上下文的接入点。

**做完的标志**：`hermes` 进对话，问一句能正常回。

### Step 4｜接辅助模型（DeepSeek，可选但建议）

```powershell
hermes config set DEEPSEEK_API_KEY sk-你的key
```

在 `hermes model` 里把"上下文压缩""会话标题"这类槽位指给 DeepSeek——这些任务量小但频次高，用便宜模型能省一大笔。

### Step 5｜验证读写本地库（关键一步）

进 `hermes` 后，在 `年轮` 目录里说：

> 「读一下 01-日常管理\日记\2026-11-22.md，用 3 条总结我今天做了什么」

再试写：

> 「把"今天确认 Hermes 能在 Windows 跑"这句话追加到 02-意识与创作\灵感\收件箱.md」

**做完的标志**：读的内容和日记一致；写的文件真的出现（去资源管理器看）。

### Step 6｜建第一个技能

Hermes 的核心能力是"从经验生成技能"。第一个技能先做允许本地写入的灵感归档：

> 「我口述灵感、创作构思和学习想法，你整理成条目送审文本；我确认后追加到 02-意识与创作\灵感\收件箱.md。饮食、作息、体重等日常记录只生成待云端 AI 处理的请求，不直接修改日记或 Excel。」

**做完的标志**：说一段测试灵感，确认后它写入正确文件；说一段饮食流水时，它不碰日常日记和台账。

### Step 7｜留积累（视频第 4 步）

```powershell
cd $env:USERPROFILE\年轮
git add . ; git commit -m "Hermes 跑通，知识库接入"
git push
```

每周日 `git log --oneline -7` 看一眼，这就是你的学习进度条。

---

## 4. 卡点自查表

| 症状 | 大概率原因 | 处理 |
|---|---|---|
| `iex : 无法加载文件` 或执行策略报错 | PowerShell 执行策略限制 | `Set-ExecutionPolicy -Scope Process Bypass` 后重跑 |
| 装完 `hermes` 不认识 | 没开新窗口 / PATH 未刷新 | 关掉重开 PowerShell；仍不行查 `%USERPROFILE%\.hermes\bin` 是否在 PATH |
| 能进对话，一提问就报错 | 没配 provider 或 key 无效 | `hermes config` 查看，`hermes model` 重选 |
| 火山 404 | 模型名填了中文名 | 换成 `ep-` 开头的接入点 ID |
| 火山 401 | key 与接入点不匹配 | 控制台重新生成，确认同一项目 |
| 报 "context window too small" | 上下文 < 64K | 换支持 128K 的接入点 |
| git clone 要求登录 | 私有仓库 | 用浏览器登录 GitHub 授权，勿贴 token |

---

## 5. 明确不做

1. 不同时装多套 agent（Hermes + Claude Code + Cursor），先跑通一个
2. 不先接 Telegram gateway，CLI 正常后再谈远程触达
3. 不追求全自动，先做到"说一句，记一条"
4. 不把 key 写进会被 git 提交的文件（`~/.hermes/.env` 在仓库外，安全）

---

*版本依据：Hermes Agent 官方文档 2026-11 版 ｜ 更新：2026-11-22（改为 Windows 路径 + 豆包主模型）*
