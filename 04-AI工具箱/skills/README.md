# Skills（AI 技能库）

这里是所有自研 skill 的**唯一权威源**。各 AI 工具（Trae、Claude Code、Cursor 等）通过**软链接**指向这里加载，改这里即全工具自动生效，无需复制同步。

## 目录约定

每个 skill 一个中文子目录（一看就懂），内含：
- `SKILL.md`：标准 skill 定义，frontmatter 里的 `name`（英文）供 AI 工具调用识别
- `install.py`：把本 skill 软链接到目标工具的 skills 目录
- 其它脚本 / 资源自包含在该目录内

| 目录 | name（英文标识） | 用途 |
|------|----------------|------|
| 八字命理 | bazi-mingli | 四柱八字双流派对照断命 |
| 八字起名 | bazi-qiming | 结合八字起名/改名（大名/小名/英文名） |
| 纳音识人 | na-yin-shi-ren | 四柱纳音识人（性格倾向与人生阶段） |
| 六爻解卦 | liu-yao-divination | 六爻摇卦解卦（吉凶/应期/建议） |
| 视频转文字到飞书 | video-transcript | 视频链接转逐字稿写飞书文档 |

## 安装到 AI 工具

单个 skill：
```bash
python 八字命理/install.py                # 装到 Trae
python 八字命理/install.py --tool claude  # 装到 Claude Code
python 八字命理/install.py --dir <目录>   # 装到自定义目录
```

全部 skill 一键装：
```bash
python install_all.py                # 全部装到 Trae
python install_all.py --tool claude  # 全部装到 Claude Code
python install_all.py --list         # 列出已识别工具目录
```

## 同步机制

- **源 = 本目录**，工具侧只是软链接（Windows 无 symlink 权限时降级 junction，junction 仅本机可用）
- 改这里的任何文件（如完善 SKILL.md），所有已链接的工具**立即生效**，零复制
- 换机器 / 换工具：clone 本仓库后跑一次 `install_all.py --tool <目标>` 即可
- 新增 skill：建一个中文子目录，放好 `SKILL.md` + `install.py`，再跑 `install_all.py`

## 迁移验收

在新目录复制任意 skill，按说明跑 `install.py` 并确认目标工具能读到 SKILL.md，即算可迁移。
