#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生活数据自动同步脚本
====================
一句话输入 → 台账5个sheet + 日记 全部同步，写后自动校验，杜绝漏更/错位/乱序。

用法（在 01-日常管理/数据 目录下执行）:
  python3 sync.py meal    --date 2026-08-20 --meal 加餐 --food "五香牛肉" --grams 100
  python3 sync.py meal    --date 2026-08-20 --meal 训练中 --food "蛋白粉(80%蛋白)" --count 1
  python3 sync.py expense --date 2026-08-20 --name "打车" --amount 13.2 --category 行
  python3 sync.py income  --date 2026-08-20 --name "闲鱼出闲置" --amount 100
  python3 sync.py sleep   --date 2026-08-20 --range "8:27-16:31"
  python3 sync.py train   --date 2026-08-20 --part "腿" --content "腿举..." --cardio 30
  python3 sync.py mantra  --date 2026-08-20 --count 108 [--note 标准]
  python3 sync.py weight  --date 2026-08-20 --kg 70.05
  python3 sync.py stock   --food "烤鹅（去皮去骨）" --consume 150
  python3 sync.py food    --name "新食物" --per100 "240,28,13,2" [--spec ...] [--unit "100g"]
  python3 sync.py deed    --date 2026-08-20 --cat 深度成长 --content "研究X并写报告" --hours 2.5
  python3 sync.py effort  --date 2026-08-20   （当日努力度评分：五维+等级+虚度标记）
  python3 sync.py effort  --week --date 2026-08-20   （一周汇总，只读）
  python3 sync.py check   --date 2026-08-20   （校验一致性，只读）
  python3 sync.py check   --all               （全量审计，只读）
  python3 sync.py today   --date 2026-08-20   （查看当日汇总，只读）

所有写命令支持 --dry-run（只打印将执行的动作，不写入）。

防错设计（对应"记错记乱决不允许"）:
  1. 写前自动备份到 备份/（保留最近30份）
  2. 先写临时文件 → 重新打开校验 → 通过才替换正式文件（原子替换）
  3. 食物/分类名称匹配：精确匹配唯一才执行，多个候选直接报错并列出（宁可问，不可猜）
  4. 每日记录/财务明细 写入后整体按日期倒序（整行读取排序，不会串行），新数据永远在顶部；
     准提咒进度保持正序（累计列才成立）；表头永远写回第1行并有内部断言
  5. 落盘顺序：台账原子写成功 → 才写日记。台账写失败时日记不动，杜绝"日记有台账无"的半写状态
  6. 每次写入后自动跑当日一致性校验并输出结果
  7. check --all 全量审计：表头逐格比对 + 三表排序 + 营养/金额/遍数跨文件一致性 + 负库存
"""

import argparse
import os
import re
import shutil
import sys
import time
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
DAILY_ROOT = os.path.dirname(BASE)
LEDGER = os.path.join(BASE, '生活数据总台账.xlsx')
DIARY_DIR = os.path.join(DAILY_ROOT, '日记')
BACKUP_DIR = os.path.join(BASE, '备份')
MONTHLY_REPORT_DIR = os.path.join(DAILY_ROOT, '报告', '月报')
TMP = LEDGER + '.tmp.xlsx'

# ============ 饮食计划目标（重组计划已启动） ============
PLAN = '重组'
TARGETS = {
    '欧包': {'train': {'kcal': 2100, 'protein': (120, 140), 'fat': 55, 'carb': 260},
             'rest': {'kcal': 1800, 'protein': (120, 140), 'fat': 55, 'carb': 0}},
    # 2026-09-23 满配四支柱定值：脂肪 55-75g（保激素）、碳水训练日 150g/休息日 115g
    # 2026-09-29 脚本对齐规范：此前误沿用压缩计划旧值（脂肪 45g、碳水 175/135）
    '重组': {'train': {'kcal': 1780, 'protein': (120, 140), 'fat': (55, 75), 'carb': 150},
             'rest': {'kcal': 1530, 'protein': (120, 140), 'fat': (55, 75), 'carb': 115}},
}

# 2026-08-08 之前的日记为旧格式（营养明细不全/格式不一），营养列以台账为准，不参与日记核对
LEGACY_UNTIL = '2026-08-07'
# 睡眠一致性校验起点（2026-11-22 立）：此前的日记作息只记主睡眠、台账记多段合计，
# 属历史格式差异非错账，故只对 2026-11-01 起的日子做逐项比对。
SLEEP_CONSIST_FROM = '2026-11-01'

EXPENSE_CATEGORIES = ['食', '住', '行', '衣', '个护', '娱乐', '🎁礼尚往来', '特殊事件', '医疗', '🎬设备', '旅行', '税费', '保险']
# 标准 emoji 形式（与台账财务明细一致，改分类请同时更新此处+台账+规范文档）
CATEGORY_EMOJI = {'食': '🍽食', '住': '🏠住', '行': '🚗行', '衣': '👔衣', '娱乐': '🥃娱乐',
                  '🎁礼尚往来': '🎁礼尚往来', '特殊事件': '💫特殊事件', '医疗': '🏥医疗',
                  '🎬设备': '🎬设备', '旅行': '✈️旅行', '税费': '🏦税费', '保险': '🛡保险',
                  '个护': '🪥个护', '🪥个护': '🪥个护'}
INCOME_DEFAULT_CAT = '💰收入'

# 表头基线：check 时逐格比对，任何错位/顶行立即可见
SHEET_HEADERS = {
    '财务明细': ['日期', '月份', '名称', '类型', '金额(¥)', '分类'],
    '每日记录': ['日期', '作息', '体重(kg)', '饮食内容', '热量(kcal)', '蛋白(g)', '脂肪(g)', '碳水(g)',
               '训练部位', '训练内容', '有氧(min)', '训练感受', '准提咒(遍)', '消费内容', '消费金额(¥)', '备注'],
    '盒马营养成分': ['食物名称', '规格', '热量(kcal/100g)', '蛋白质(g/100g)', '脂肪(g/100g)', '碳水(g/100g)',
                 '常用分量', '分量热量(kcal)', '数据来源', '备注'],
    '准提咒进度': ['日期', '遍数', '备注', '累计天数', '累计遍数', '目标', '完成度'],
    '食物库存明细': ['食材名称', '规格', '当前库存', '单位', '库存状态', '最近更新日期', '备注'],
    '成长流水': ['日期', '类别', '领域', '性质', '内容', '时长(h)', '产出/成果', '备注'],
    '每日努力度': ['日期', '深度成长(h)', '事务(h)', '训练状态', '准提咒(遍)', '睡眠(h)', '管理分', '综合分', '完整度', '等级', '虚度标记'],
    '体成分记录': ['日期', '体重(kg)', 'BMI', '体脂率(%)', '肌肉量(kg)', '肌肉率(%)', '骨骼肌量(kg)', '骨骼肌率(%)',
               '骨量(kg)', '水分率(%)', '蛋白质率(%)', '内脏脂肪(级)', '皮下脂肪率(%)', '去脂体重(kg)',
               '基础代谢(kcal)', '身体年龄(岁)', '备注'],
    '身材维度记录': ['日期', '胸围(cm)', '腰围(cm)', '臀围(cm)', '手臂围(cm)', '大腿围(cm)', '小腿围(cm)',
                '腰臀比', '胸腰比', '备注'],
    '健身成本记录': ['日期', '是否训练', '训练部位', '有氧(min)', '累计训练天数', '累计健身支出(¥)', '单次成本(¥/次)', '备注'],
}

# 体成分记录表第3列起的14个指标（weight --body 的 CSV 顺序，与表头一一对应）
# 2026-09-06 删除"身体得分"项（用户裁定无参考价值）：原第17列已删，备注由第18列左移至第17列
BODY_COLS = ['BMI', '体脂率', '肌肉量', '肌肉率', '骨骼肌量', '骨骼肌率', '骨量', '水分率', '蛋白质率',
             '内脏脂肪', '皮下脂肪率', '去脂体重', '基础代谢', '身体年龄']

MEAL_ORDER = ['起床餐', '训练前', '练前餐', '训练中补给', '训练中', '午餐', '加餐', '练后加餐', '晚餐', '夜宵', '饮品']

# ============ 努力度评估（恐惧驱动·客观标准，2026-08-20 建立） ============
# 研究锚点（已查证）：
#   精英刻意练习：柏林音乐学院最佳组小提琴手日均约3.5h（日记周24.3h），人类上限约4h（Ericsson 1993；Cal Newport《Deep Work》）
#   普通人真实产出：办公室职员约2h53m/天（Vouchercloud约2000人调查；Cognition期刊2h48m）
#   精英恢复：最佳组平均睡8.6h并午睡——睡眠是产出的一部分，不是懈怠
DEED_CATS = ['深度成长', '事务', '训练', '修行', '恢复', '其他']
# 领域=精力资产池（只对"深度成长"有意义）。新领域必须先问用户确认 → 确认后加进此清单。
KNOWN_DOMAINS = ['剧本创作', '系统建设', '六爻研究', '视频创作', '八字研究', '跨境电商']
IO_TYPES = ['输入', '输出']
OUTPUT_RATIO_FLOOR = 0.30  # 输出占比健康线（暂定，输出h/深度成长总h ≥ 30%）
DEPTH_ANCHORS = [(0, 0), (1, 20), (2.88, 60), (3.5, 85), (4, 100)]  # h→分：0 / 新手1h / 普通人2h53 / 精英日均3.5 / 天花板4h
SLEEP_ANCHORS = [(5, 30), (6, 55), (7, 90), (8, 100), (9, 90), (10, 70), (11, 50)]
EFFORT_WEIGHTS = {'深度成长': 0.60, '训练': 0.20, '睡眠': 0.10, '自我管理': 0.10}  # 基础项 100 分
EFFORT_BONUS = {'准提咒': 5.0}  # 附加项：做了加分，不做不扣，也不计入完整度
GRADES = [(90, 'S 精英日'), (80, 'A 两成日'), (70, 'B 中坚日'), (60, 'C 普通日'), (0, 'D 虚度警报')]

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402


class Stop(Exception):
    """宁可问，不可猜：抛出即中止，绝不带猜测写入。"""


def norm_date(s):
    return str(s).strip()[:10]


def norm_name(s):
    """去掉品牌修饰词后比较，用于模糊但必须唯一的匹配。"""
    s = str(s or '').strip()
    for token in ['盒马工坊', '盒马日日鲜', '盒马刺身', '盒马', '（杨生记）', '(杨生记)', '(盒马工坊HPP)', '(HPP锁鲜)']:
        s = s.replace(token, '')
    return re.sub(r'[\s（）()]', '', s)


# ==================== 台账读写 ====================

def load():
    return openpyxl.load_workbook(LEDGER)


def sheet_rows(ws, ncol=None):
    """整行读取为元组列表（保序），用于安全排序。"""
    ncol = ncol or ws.max_column
    return [tuple(ws.cell(r, c).value for c in range(1, ncol + 1)) for r in range(1, ws.max_row + 1)]


def sort_by_date(ws, desc=True, date_col=1, ncol=None):
    """整行排序：表头永远写回第1行；日期行紧随其后（按方向），无日期行垫底。
    稳定排序：同日期的行保持写入时的相对顺序（insert_rows(2) 的新行在同日期中置顶）。
    写回前必须先显式清空整行再写：openpyxl 的 cell(r,c,None) 对 None 参数不写入，
    旧行残值会留在原位造成跨行污染（2026-09-06 体成分对角线残迹的根因，
    已实证复现：9/5 行继承 9/4 行的得分/备注）。"""
    rows = sheet_rows(ws, ncol)
    if not rows or rows[0][0] is None:
        raise Stop(f'[{ws.title}] 第1行表头为空，结构异常，拒绝排序写入。')
    header, body = rows[0], rows[1:]
    dated, undated = [], []
    for row in body:
        d = norm_date(row[date_col - 1])
        (dated if re.match(r'^\d{4}-\d{2}-\d{2}', d) else undated).append(row)
    dated.sort(key=lambda r: norm_date(r[date_col - 1]), reverse=desc)
    ordered = [header] + dated + undated
    if len(ordered) != len(rows):
        raise Stop(f'[{ws.title}] 排序前后行数不一致 {len(rows)}→{len(ordered)}，拒绝写入。')
    width = len(header)
    for i, row in enumerate(ordered, start=1):
        for c in range(1, width + 1):
            ws.cell(i, c).value = None   # 显式清空（属性赋值才生效）
        for c, v in enumerate(row, start=1):
            if v is not None:
                ws.cell(i, c).value = v  # 仅写非空值
    if ws.cell(1, 1).value != header[0]:
        raise Stop(f'[{ws.title}][内部断言] 表头未正确写回，放弃本次写入。')
    return ws


def save_atomic(wb, checks=()):
    """备份 → 写临时文件 → 重开校验 → 原子替换。任何一步失败都不碰正式文件。"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(BACKUP_DIR, f'生活数据总台账_{stamp}.xlsx')
    shutil.copy2(LEDGER, backup)
    wb.save(TMP)
    wb2 = openpyxl.load_workbook(TMP)
    for sheet, row, col, expect, label in checks:
        got = wb2[sheet].cell(row, col).value
        if got != expect:
            raise Stop(f'校验失败[{label}] 期望 {expect!r} 实得 {got!r}，已放弃写入，正式文件未动。备份：{backup}')
    os.replace(TMP, LEDGER)
    # 只保留最近30份备份
    backups = sorted(f for f in os.listdir(BACKUP_DIR) if f.endswith('.xlsx'))
    for old in backups[:-30]:
        os.remove(os.path.join(BACKUP_DIR, old))
    return backup


# ==================== 定位工具 ====================

def find_daily_row(ws, date):
    for r in range(2, ws.max_row + 1):
        if norm_date(ws.cell(r, 1).value) == date:
            return r
    raise Stop(f'每日记录中没有 {date} 的行，请先确认日期。')


def ensure_daily_row(ws, date):
    """没有当日行则新建（倒序表，新行插在第2行）。"""
    try:
        return find_daily_row(ws, date), False
    except Stop:
        ws.insert_rows(2)
        ws.cell(2, 1, date)
        return 2, True


def match_food(ws, name):
    """营养成分表匹配：精确→规范化唯一。多候选/零候选都报错。"""
    exact = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)
             if str(ws.cell(r, 1).value or '').strip() == name.strip()]
    if len(exact) == 1:
        return exact[0][0]
    if len(exact) > 1:
        raise Stop(f'营养成分表中有 {len(exact)} 条同名"{name}"，请先去重：{[n for _, n in exact]}')
    cand = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)
            if norm_name(ws.cell(r, 1).value) == norm_name(name)]
    if len(cand) == 1:
        return cand[0][0]
    if not cand:
        cand = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)
                if norm_name(name) in norm_name(ws.cell(r, 1).value)
                or norm_name(ws.cell(r, 1).value) in norm_name(name)]
    if len(cand) == 1:
        print(f'⚠️ 按名称匹配到：{cand[0][1]}（非精确名，请确认）')
        return cand[0][0]
    if not cand:
        raise Stop(f'营养成分表中没有"{name}"。请先用 food 命令录入，或告诉我营养数据。')
    raise Stop(f'"{name}" 匹配到 {len(cand)} 条，请说清是哪一种：\n  ' + '\n  '.join(n for _, n in cand))


def match_stock(ws, name):
    # 营养成分表名 → 库存表名 别名映射（2026-10-04 立：烤鸡名不匹配曾 3 次静默漏扣）
    STOCK_ALIAS = {'烤鸡(去骨带皮·估)': '烤鸡(去骨·冷藏)'}
    name = STOCK_ALIAS.get(name.strip(), name)
    exact = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)
             if str(ws.cell(r, 1).value or '').strip() == name.strip()]
    if len(exact) == 1:
        return exact[0][0]
    cand = [(r, ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)
            if norm_name(name) in norm_name(ws.cell(r, 1).value)
            or norm_name(ws.cell(r, 1).value) in norm_name(name)]
    if len(cand) == 1:
        return cand[0][0]
    if not cand:
        return None  # 不在库存跟踪范围内的食物，允许跳过
    raise Stop(f'库存表"{name}"匹配到多条，请说清是哪一种：\n  ' + '\n  '.join(n for _, n in cand))


def match_category(cat):
    c = str(cat).strip()
    for std in EXPENSE_CATEGORIES:
        if c == std or c in std or std in c:
            return CATEGORY_EMOJI.get(std, std)
    raise Stop(f'分类"{cat}"不在标准分类中：{"、".join(EXPENSE_CATEGORIES)}')


def parse_unit_grams(text):
    """从"40g/个" "100g(1盒)" "30g(1勺)" "25mL(1条)" 解析单份克数。"""
    m = re.search(r'([\d.]+)\s*[gm][lr]?', str(text), re.I)
    return float(m.group(1)) if m else None


def guess_unit_grams(text):
    """从"40g×8个" "175g/5枚(35g/枚)" "950ml/盒" 等规格解析单份克数（取最小g数字）。"""
    nums = [float(x) for x in re.findall(r'([\d.]+)\s*[gm][lr]?', str(text), re.I)]
    return min(nums) if nums else None


def is_liquid(spec, common_portion=''):
    """判断食物是否为液体（规格或常用分量中含 ml/ML/mL 等）。"""
    text = f'{spec or ""} {common_portion or ""}'
    return bool(re.search(r'm[lL]', text))


# ==================== 营养计算 ====================

def food_macros(ws, row, grams):
    per100 = [float(ws.cell(row, c).value) for c in (3, 4, 5, 6)]
    return [round(v * grams / 100, 1) for v in per100]


# ==================== 日记读写 ====================

DIARY_TEMPLATE = """# {date} 每日记录

## 作息
- 待记录

## 时间轴

| 时段 | 时长 | 类别 | 内容 | 标记 |
|------|------|------|------|------|
### 时间未知
- （暂无）

## 准提咒
- 待记录

## 训练
### 今日计划：待确认

### 实际训练记录
- 待记录

## 成长记录（今日做了什么）
- 待记录

## 努力度评估
- 待评估

## 饮食
### 今日营养汇总
| 营养素 | 目标 | 已摄入 | 剩余 | 状态 |
|--------|------|--------|------|------|
| 热量 | 0 kcal | 0 kcal | 0 kcal | 0% |
| 蛋白质 | 120-140g | 0g | 还需120g | 🥩 不够 |
| 脂肪 | 55-75g | 0g | 还差55g | ⚠️ 不够 |
| 碳水 | 150g | 0g | 剩150g | ✅ 正常 |

## 消费

## 库存更新

## 备注
"""


def diary_path(date):
    return os.path.join(DIARY_DIR, f'{date}.md')


def load_diary(date):
    p = diary_path(date)
    if not os.path.exists(p):
        with open(p, 'w', encoding='utf-8') as f:
            f.write(DIARY_TEMPLATE.format(date=date))
        print(f'（已创建当日日记 {date}.md）')
    with open(p, encoding='utf-8') as f:
        return f.read()


def save_diary(date, text):
    """写日记 + 防回滚校验（2026-09-02 加）。

    环境偶发"写后文件回退到旧版本"（2026-09-02 当天两度发生，曾丢失整段时间轴/加餐记录）。
    对策：写后延时 0.4s 回读全文比对，不一致即重写，最多 5 次；仍失败则抛错让操作者
    察觉（此时台账可能已写，是半写状态，显式报错好过静默丢数据）。所有写日记的
    命令（meal/expense/sleep/train/mantra/timeline…）自动获得此保护。
    """
    p = diary_path(date)
    for attempt in range(1, 6):
        with open(p, 'w', encoding='utf-8') as f:
            f.write(text)
        time.sleep(0.4)
        with open(p, encoding='utf-8') as f:
            if f.read() == text:
                if attempt > 1:
                    print(f'✅ 日记第{attempt}次写入后校验通过')
                return
        print(f'⚠️ 日记写入后被回滚（第{attempt}次），重试')
    raise Stop(f'日记 {date} 写入 5 次均被回滚——台账可能已写而日记未写（半写状态），请立即检查。')


def fmt_dur(hours):
    m = round(hours * 60)
    return f'{m // 60}h{m % 60:02d}m'


def parse_range(rng):
    """'8:27-16:31'（支持全角）→ 小时数。跨零点自动+24h。"""
    s = rng.replace('：', ':').replace('—', '-').replace('～', '-').strip()
    m = re.match(r'^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$', s)
    if not m:
        raise Stop(f'睡眠时间段格式看不懂：{rng}（示例：8:27-16:31）')
    h1, m1, h2, m2 = map(int, m.groups())
    start, end = h1 + m1 / 60, h2 + m2 / 60
    if end <= start:
        end += 24
    return start, end, f'{h1}:{m1:02d}→{h2}:{m2:02d}'


# ---------- 日记：饮食 ----------

def diary_add_meal(text, meal, display, macros):
    """在对应餐次表格插入食物行并重算该表合计；没有该餐次则新建。"""
    kcal, p, f_, c = macros
    food_line = f'| {display} | — | {kcal} | {p} | {f_} | {c} |'
    lines = text.split('\n')
    head_idx = next((i for i, l in enumerate(lines) if l.startswith(f'### {meal}')), None)
    if head_idx is not None:
        # 找该节的合计行
        total_idx = next((i for i in range(head_idx + 1, len(lines))
                          if lines[i].startswith('| **合计**') or (lines[i].startswith('## ') and i > head_idx)), None)
        if total_idx is None or not lines[total_idx].startswith('| **合计**'):
            raise Stop(f'日记"{meal}"小节结构异常（找不到合计行），请手动检查，不自动写入。')
        lines.insert(total_idx, food_line)
        # 重算该节合计
        vals = [0.0] * 4
        for i in range(head_idx + 1, total_idx + 1):
            row = [c.strip() for c in lines[i].strip('|').split('|')] if lines[i].startswith('|') else []
            if len(row) == 6 and not row[0].startswith('**') and not row[0].startswith('食物'):
                try:
                    vals = [round(a + float(re.sub(r'[^\d.\-]', '', x) or 0), 1)
                            for a, x in zip(vals, row[2:6])]
                except ValueError:
                    pass
        lines[total_idx + 1] = f'| **合计** | — | **{vals[0]:.1f}** | **{vals[1]:.1f}** | **{vals[2]:.1f}** | **{vals[3]:.1f}** |'
    else:
        # 新建餐次小节：插在"### 今日营养汇总"前
        anchor = next((i for i, l in enumerate(lines) if l.startswith('### 今日营养汇总')), None)
        if anchor is None:
            raise Stop('日记缺少"### 今日营养汇总"锚点，不自动新建餐次，请手动检查。')
        new = [f'### {meal}', '| 食物 | 重量 | 热量(kcal) | 蛋白(g) | 脂肪(g) | 碳水(g) |',
               '|------|------|-----------|---------|---------|---------|',
               food_line,
               f'| **合计** | — | **{kcal}** | **{p}** | **{f_}** | **{c}** |', '']
        lines[anchor:anchor] = new
    return '\n'.join(lines)


def diary_totals(text):
    """汇总饮食区所有餐次表格的食物行营养。兼容新旧两种日记格式：
    新：'## 饮食'；旧：'## 实际饮食记录'。表格允许第7列备注。"""
    lines = text.split('\n')
    start = next((i for i, l in enumerate(lines)
                  if l.startswith('## 饮食') or l.startswith('## 实际饮食记录')), None)
    if start is None:
        return None  # 无饮食区，视为不可解析（不当作0）
    end = next((i for i, l in enumerate(lines) if i > start and l.startswith('## ')), len(lines))
    totals = [0.0] * 4
    for line in lines[start:end]:
        if not line.startswith('|'):
            continue
        row = [c.strip() for c in line.strip('|').split('|')]
        if len(row) in (6, 7) and not row[0].startswith('**') and not row[0].startswith('食物') and not row[0].startswith('-'):
            nums = []
            for x in row[2:6]:
                m = re.search(r'-?\d+\.?\d*', x)
                nums.append(float(m.group()) if m else 0.0)
            totals = [round(a + b, 1) for a, b in zip(totals, nums)]
    return totals


def diary_refresh_summary(text, train_day=True):
    """按当前计划目标重写"今日营养汇总"四行。"""
    t = TARGETS[PLAN]['train' if train_day else 'rest']
    kcal, p, f_, c = diary_totals(text)
    rows = []
    remain = t['kcal'] - kcal
    pct = round(kcal / t['kcal'] * 100) if t['kcal'] else 0
    rows.append(f'| 热量 | {t["kcal"]} kcal | {kcal:.1f} kcal | '
                f'{"剩" if remain >= 0 else "超"}{abs(remain):.1f} kcal | {pct}% |')
    lo, hi = t['protein']
    if p < lo:
        rows.append(f'| 蛋白质 | {lo}-{hi}g | {p:.1f}g | 还需{lo - p:.1f}g | 🥩 不够 |')
    elif p > hi:
        rows.append(f'| 蛋白质 | {lo}-{hi}g | {p:.1f}g | 超{p - hi:.1f}g | ✅ 超标 |')
    else:
        rows.append(f'| 蛋白质 | {lo}-{hi}g | {p:.1f}g | 区间内 | ✅ 达标 |')
    flo, fhi = t['fat']
    if f_ < flo:
        rows.append(f'| 脂肪 | {flo}-{fhi}g | {f_:.1f}g | 还差{flo - f_:.1f}g | ⚠️ 不够 |')
    elif f_ > fhi:
        rows.append(f'| 脂肪 | {flo}-{fhi}g | {f_:.1f}g | 超{f_ - fhi:.1f}g | ✅ 超标 |')
    else:
        rows.append(f'| 脂肪 | {flo}-{fhi}g | {f_:.1f}g | 区间内 | ✅ 达标 |')
    if t['carb']:
        rows.append(f'| 碳水 | {t["carb"]}g | {c:.1f}g | '
                    f'{"剩" if t["carb"] >= c else "超"}{abs(t["carb"] - c):.1f}g | '
                    f'{"✅ 正常" if t["carb"] >= c else "🔴 超标"} |')
    else:
        rows.append(f'| 碳水 | 不限 | {c:.1f}g | — | — |')
    lines = text.split('\n')
    anchor = next((i for i, l in enumerate(lines) if l.startswith('### 今日营养汇总')), None)
    if anchor is None:
        raise Stop('日记缺少"### 今日营养汇总"，请手动检查。')
    # 替换锚点后直到下一个 ### / ## 的4行数据
    i = anchor + 1
    while i < len(lines) and not lines[i].startswith('#'):
        if lines[i].startswith('| 热量') or lines[i].startswith('| 蛋白质') \
           or lines[i].startswith('| 脂肪') or lines[i].startswith('| 碳水'):
            lines[i] = rows.pop(0) if rows else lines[i]
        i += 1
    return '\n'.join(lines)


# ---------- 日记：消费/睡眠/准提咒/体重/库存 ----------

def diary_add_expense(text, name, amount, cat):
    lines = text.split('\n')
    for i, l in enumerate(lines):
        if l.startswith('## 消费'):
            j = i + 1
            while j < len(lines) and not lines[j].startswith('## '):
                if lines[j].startswith('- **合计'):
                    lines.insert(j, f'- {name}：¥{amount:.2f}（{cat}）')
                    return '\n'.join(lines)
                j += 1
            lines.insert(i + 1, f'- {name}：¥{amount:.2f}（{cat}）')
            return '\n'.join(lines)
    raise Stop('日记缺少"## 消费"小节，请手动检查。')


def diary_expense_total(text):
    """逐项求和（排除合计行自身与收入行），并返回日记合计行的值。
    兼容两种历史格式：'- **合计**：¥x' 与 '- **合计：¥x**'。"""
    total = 0.0
    for line in text.split('\n'):
        if re.match(r'^- \*\*合计', line):
            continue
        m = re.match(r'^- (?!💰收入)(.+?)：¥(-?[\d.]+)', line)
        if m:
            total += float(m.group(2))
    for line in text.split('\n'):
        m = re.search(r'\*\*合计\*\*[：:]¥(-?[\d.]+)', line) \
            or re.search(r'\*\*合计[：:]¥(-?[\d.]+)\*\*', line)
        if m:
            return round(total, 2), float(m.group(1))
    return round(total, 2), None


def diary_set_expense_total(text, total):
    """重写合计行，统一为 '- **合计：¥x.xx**'（用户日记现行格式）。"""
    new = f'- **合计：¥{total:.2f}**'
    text = re.sub(r'^- \*\*合计\*\*[：:]¥-?[\d.]+', new, text, count=1, flags=re.M)
    text = re.sub(r'^- \*\*合计[：:]¥-?[\d.]+\*\*', new, text, count=1, flags=re.M)
    return text


def diary_add_income(text, name, amount, cat):
    lines = text.split('\n')
    for i, l in enumerate(lines):
        if l.startswith('## 消费'):
            lines.insert(i + 1, f'- 💰收入｜{name}：+¥{amount:.2f}（{cat}）')
            return '\n'.join(lines)
    raise Stop('日记缺少"## 消费"小节，请手动检查。')


def diary_add_sleep(text, label, rng):
    _, _, arrow = parse_range(rng)
    start, end, _ = parse_range(rng)
    lines = text.split('\n')
    line = f'- {label}：{arrow.replace("→", " 入睡 → ")} 起床（约 {fmt_dur(end - start)}）'
    for i, l in enumerate(lines):
        if l.startswith('## 作息'):
            if '- 待记录' in lines[i + 1]:
                lines[i + 1] = line
            else:
                lines.insert(i + 1, line)
            return '\n'.join(lines)
    raise Stop('日记缺少"## 作息"小节。')


def diary_sleep_total(text):
    """作息小节里所有"约 XhXXm"求和。"""
    start = next((i for i, l in enumerate(text.split('\n')) if l.startswith('## 作息')), 0)
    end = next((i for i, l in enumerate(text.split('\n')) if i > start and l.startswith('## ')), 10 ** 9)
    mins = 0
    for line in text.split('\n')[start:end]:
        for h, m in re.findall(r'约\s*(\d+)h(\d{1,2})m?', line):
            mins += int(h) * 60 + int(m)
    return mins


def diary_set_mantra(text, count, note, cum, goal):
    lines = text.split('\n')
    line = f'- {count}遍（{note}）· 累计 {cum} 遍 / 目标 {goal} 遍（{cum / goal * 100:.1f}%）'
    for i, l in enumerate(lines):
        if l.startswith('## 准提咒'):
            if i + 1 < len(lines) and '- 待记录' in lines[i + 1]:
                lines[i + 1] = line
            else:
                lines.insert(i + 1, line)
            return '\n'.join(lines)
    raise Stop('日记缺少"## 准提咒"小节。')


def diary_add_stock(text, line):
    if not line.startswith('- '):
        line = '- ' + line  # 与日记其他列表行格式统一
    lines = text.split('\n')
    for i, l in enumerate(lines):
        if l.startswith('## 库存更新'):
            lines.insert(i + 1, line)
            return '\n'.join(lines)
    raise Stop('日记缺少"## 库存更新"小节。')


def diary_add_weight(text, kg, extra=''):
    lines = text.split('\n')
    for i, l in enumerate(lines):
        if l.startswith('## 备注'):
            lines.insert(i + 1, f'- ⚖️ 称重：{kg}kg{extra}')
            return '\n'.join(lines)
    raise Stop('日记缺少"## 备注"小节。')


def cmd_dim(a):
    """身材维度记录：胸/腰/臀/臂/大腿/小腿，自动算腰臀比+胸腰比"""
    vals = [a.chest, a.waist, a.hip, a.arm, a.thigh, a.calf]
    if any(v <= 0 for v in vals):
        raise Stop('六个围度都必须为正数（cm）。')
    whr = round(a.waist / a.hip, 3)
    cwr = round(a.chest / a.waist, 2)
    print(f'→ 维度：胸{a.chest} 腰{a.waist} 臀{a.hip} 臂{a.arm} 大腿{a.thigh} 小腿{a.calf} | 腰臀比{whr} 胸腰比{cwr}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    dws = wb['身材维度记录']
    r = next((i for i in range(2, dws.max_row + 1) if norm_date(dws.cell(i, 1).value) == a.date), None)
    if r is None:
        dws.insert_rows(2)
        dws.cell(2, 1, a.date)
        r = 2
    for c, v in enumerate([a.chest, a.waist, a.hip, a.arm, a.thigh, a.calf, whr, cwr], start=2):
        dws.cell(r, c, v)
    if a.note:
        dws.cell(r, 10, a.note)
    sort_by_date(dws, desc=True, ncol=len(SHEET_HEADERS['身材维度记录']))
    r = next(i for i in range(2, dws.max_row + 1) if norm_date(dws.cell(i, 1).value) == a.date)
    checks = [('身材维度记录', r, 7, a.calf, '小腿围同步'), ('身材维度记录', r, 8, whr, '腰臀比同步')]
    dim_line = f'- 📏 胸{a.chest} / 腰{a.waist} / 臀{a.hip} / 臂{a.arm} / 大腿{a.thigh} / 小腿{a.calf} cm（腰臀比 {whr}）'
    text = load_diary(a.date)
    lines = text.split('\n')
    if any(l.startswith('## 身材维度') for l in lines):
        for i, l in enumerate(lines):
            if l.startswith('## 身材维度'):
                lines.insert(i + 1, dim_line)
                break
    else:
        for i, l in enumerate(lines):
            if l.startswith('## 备注'):
                lines[i:i] = ['## 身材维度', dim_line, '']
                break
        else:
            raise Stop('日记缺少"## 备注"小节。')
    save_atomic(wb, checks)
    save_diary(a.date, '\n'.join(lines))
    print('✅ 已同步：身材维度记录 + 日记')


# ==================== 各业务命令 ====================

def cmd_meal(a):
    wb = load()
    nws = wb['盒马营养成分']
    row = match_food(nws, a.food)
    name = nws.cell(row, 1).value
    unit_g = None
    if a.grams:
        grams = a.grams
    elif a.count:
        unit_g = parse_unit_grams(nws.cell(row, 7).value) or guess_unit_grams(nws.cell(row, 2).value)
        if not unit_g:
            raise Stop(f'"{name}"解析不出单份克数（常用分量={nws.cell(row, 7).value}），请用 --grams 直接给克数。')
        grams = round(a.count * unit_g, 1)
    else:
        raise Stop('请给 --grams 克数 或 --count 个数（袋/盒/条/勺）。')
    macros = food_macros(nws, row, grams)
    spec = nws.cell(row, 2).value or ''
    common = nws.cell(row, 7).value or ''
    liquid = is_liquid(spec, common)
    if a.count:
        display = a.display or f'{name} {a.count}个'
    elif liquid:
        display = a.display or f'{name} {grams:g}ml'
    else:
        display = a.display or f'{name} {grams:g}g'
    print(f'→ {display}：{macros[0]} kcal / 蛋白{macros[1]} / 脂肪{macros[2]} / 碳水{macros[3]}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return

    date = a.date
    # 1) 日记
    text = load_diary(date)
    text = diary_add_meal(text, a.meal, display, macros)
    dws = wb['每日记录']
    r, _ = ensure_daily_row(dws, date)
    train_day = '休息' not in str(dws.cell(r, 9).value or '')
    text = diary_refresh_summary(text, train_day)
    # 2) 每日记录（营养列以日记全表求和回写：与校验同源，手工合并日记行后自愈，杜绝两路累加漂移）
    totals = diary_totals(text)
    if totals is None:
        raise Stop('日记里找不到饮食区（## 饮食 / ## 实际饮食记录），请手动检查。')
    old = str(dws.cell(r, 4).value or '').strip()
    if old and old not in ('', '待记录'):
        dws.cell(r, 4, f'{old} + {a.meal}:{display.split(" ")[0]}({macros[0]:.0f})')
    else:
        dws.cell(r, 4, f'{a.meal}:{display.split(" ")[0]}({macros[0]:.0f})')
    for c, v in zip((5, 6, 7, 8), (round(t, 1) for t in totals)):
        dws.cell(r, c, v)
    # 3) 库存
    stock_note = ''
    sws = wb['食物库存明细']
    srow = match_stock(sws, name)
    if srow:
        unit = str(sws.cell(srow, 4).value or '')
        cur = sws.cell(srow, 3).value
        if unit in ('g', 'ml') and isinstance(cur, (int, float)):
            after = round(cur - grams, 1)
            sws.cell(srow, 3, after)
            sws.cell(srow, 5, '⚠️ 已吃完' if after <= 0 else '✅ 充足')
            sws.cell(srow, 6, date)
            stock_note = f'{name}：{after}{unit} ← {cur} - {grams}{unit}（{a.meal}）'
        elif a.count and isinstance(cur, (int, float)):
            # 计数型单位防呆：库存>50 或 单次扣减>20 疑似异常，拒绝操作（防止636个欧包这类事故）
            if any(unit.startswith(u) for u in ('个', '袋', '盒', '枚', '份', '根', '瓶', '条')):
                if cur > 50 or a.count > 20:
                    raise Stop(f'⛔ "{name}"单位是"{unit}"（计数型），当前库存 {cur}{unit} / 本次扣减 {a.count}{unit}，'
                               f'数值明显异常，疑似克数与个数混淆。请用 stock --set 先修正库存后再扣。')
            after = round(cur - a.count, 1)
            sws.cell(srow, 3, after)
            sws.cell(srow, 5, '⚠️ 已吃完' if after <= 0 else '✅ 充足')
            sws.cell(srow, 6, date)
            stock_note = f'{name}：{after}{unit} ← {cur} - {a.count}{unit}（{a.meal}）'
        else:
            stock_note = f'⚠️ {name}在库存表中但单位是"{unit}"（库存{cur}），无法按克自动扣减，请确认。'
    else:
        # 匹配失败必须留痕（2026-10-04 立：烤鸡名不匹配曾 3 次静默漏扣库存）
        print(f'⚠️ 库存表未匹配到「{name}」，本次未扣库存——若该食材在库存管理内，请核对条目名或告诉我加别名')
    if stock_note:
        text = diary_add_stock(text, stock_note)
        print(f'→ 库存：{stock_note}')
    # 4) 校验（营养列已按同源求和写入，此处防写入落空/错位）
    kcal, p, f_, c = totals
    checks = [('每日记录', r, 5, round(kcal, 1), '热量同步'),
              ('每日记录', r, 6, round(p, 1), '蛋白同步')]
    save_atomic(wb, checks)  # 台账原子写成功后才落盘日记，杜绝"日记有台账无"的半写状态
    save_diary(date, text)
    print(f'✅ 已同步：日记 + 每日记录 + 库存 | 今日累计 {kcal:.1f} kcal / 蛋白{p:.1f}g / 脂肪{f_:.1f}g / 碳水{c:.1f}g')


def cmd_expense(a):
    cat = match_category(a.category)
    if a.amount == 0:
        raise Stop('金额不能为 0。')
    date = a.date
    wb = load()
    fws = wb['财务明细']
    name = a.name if a.amount > 0 else (a.name if a.name.startswith('退款') else f'退款-{a.name}')
    amount = a.amount
    print(f'→ 财务明细：{date} | {name} | ¥{amount:.2f} | {cat}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    fws.insert_rows(2)
    for c, v in enumerate([date, date[:7], name, 'expense', round(amount, 2), cat], 1):
        fws.cell(2, c, v)
    sort_by_date(fws, desc=True)
    dws = wb['每日记录']
    r, _ = ensure_daily_row(dws, date)
    old = str(dws.cell(r, 14).value or '').strip()
    dws.cell(r, 14, f'{old} + {name}¥{amount:.2f}' if old else f'{name}¥{amount:.2f}')
    dws.cell(r, 15, round((dws.cell(r, 15).value or 0) + amount, 2))
    text = load_diary(date)
    text = diary_add_expense(text, name, amount, cat)
    total, _ = diary_expense_total(text)
    text = diary_set_expense_total(text, total)
    row_of_new = next(rr for rr in range(2, fws.max_row + 1) if norm_date(fws.cell(rr, 1).value) == date
                      and fws.cell(rr, 3).value == name and fws.cell(rr, 5).value == round(amount, 2))
    save_atomic(wb, [('财务明细', row_of_new, 3, name, '财务行存在'),
                     ('每日记录', r, 15, round(total, 2), '消费金额一致')])
    save_diary(date, text)
    print(f'✅ 已同步：财务明细(顶部) + 每日记录 + 日记 | 当日支出合计 ¥{total:.2f}')


def cmd_income(a):
    date = a.date
    cat = a.category or INCOME_DEFAULT_CAT
    print(f'→ 财务明细(收入)：{date} | {a.name} | +¥{a.amount:.2f} | {cat}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    fws = wb['财务明细']
    fws.insert_rows(2)
    for c, v in enumerate([date, date[:7], a.name, 'income', round(a.amount, 2), cat], 1):
        fws.cell(2, c, v)
    sort_by_date(fws, desc=True)
    text = load_diary(date)
    text = diary_add_income(text, a.name, a.amount, cat)
    row_of_new = next(rr for rr in range(2, fws.max_row + 1) if norm_date(fws.cell(rr, 1).value) == date
                      and fws.cell(rr, 3).value == a.name)
    save_atomic(wb, [('财务明细', row_of_new, 4, 'income', '收入类型')])
    save_diary(date, text)
    print(f'✅ 已同步：财务明细(income) + 日记 | 收入不计入每日支出列')


def cmd_sleep(a):
    start, end, _ = parse_range(a.range)
    date = a.date
    print(f'→ 睡眠：{a.range}（{fmt_dur(end - start)}）')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    text = load_diary(date)
    n = len(re.findall(r'^- .+入睡', text.split('## 训练')[0].split('## 作息')[-1], flags=re.M))
    label = a.label or ('主睡眠' if n == 0 else f'小睡{n}')
    text = diary_add_sleep(text, label, a.range)
    mins = diary_sleep_total(text)
    summary = f'{label} {a.range.replace("-", "→")}({fmt_dur(end - start)})' if label == '主睡眠' \
        else f'{mins // 60}h{mins % 60:02d}m（{n + 1}段）'
    if label != '主睡眠':
        # 多段：压缩为 "段1+段2=总"（箭头前可能带"入睡/（8/28 晚）"等说明，需容忍）
        segs = re.findall(r'^- (.+?)：(\d{1,2}:\d{2}).*?→\s*(\d{1,2}:\d{2})', text, flags=re.M)
        if segs:
            summary = '+'.join(f'{s[1]}-{s[2]}' for s in segs) + f'={mins // 60}h{mins % 60:02d}m'
        else:
            # 兜底：不写以 = 开头的串，避免被 Excel 当成公式
            summary = fmt_dur(mins / 60)
    wb = load()
    dws = wb['每日记录']
    r, _ = ensure_daily_row(dws, date)
    dws.cell(r, 2, summary)
    save_atomic(wb, [('每日记录', r, 2, summary, '作息同步')])
    save_diary(date, text)
    print(f'✅ 已同步：日记作息 + 每日记录 | 当日睡眠合计 {mins // 60}h{mins % 60:02d}m')


# ---------- 时间轴（2026-09-02 新增：补录一行/闭合开放行，替代临时手写脚本省对话轮次） ----------

def fmt_timeline_dur(mins):
    """分钟数 → 日记时间轴时长风格：36min / 2h / 2h28m。"""
    if mins <= 0:
        return '—'
    h, m = mins // 60, mins % 60
    if h == 0:
        return f'{m}min'
    return f'{h}h' if m == 0 else f'{h}h{m:02d}m'


def parse_hhmm(s, what):
    m = re.match(r'^(\d{1,2}):(\d{2})$', str(s).strip())
    if not m or not (0 <= int(m.group(1)) <= 23 and 0 <= int(m.group(2)) <= 59):
        raise Stop(f'{what}格式看不懂：{s}（示例：16:40；小时0-23，分钟0-59）')
    return int(m.group(1)) * 60 + int(m.group(2))


def cmd_timeline(a):
    date = a.date
    text = load_diary(date)
    lines = text.split('\n')
    head = next((i for i, l in enumerate(lines) if l.startswith('## 时间轴')), None)
    if head is None:
        raise Stop('日记缺少"## 时间轴"锚点，请手动检查。')
    data_rows = [i for i in range(head + 1, len(lines)) if re.match(r'^\|\s*\d{1,2}:\d{2}', lines[i])]
    if not data_rows:
        raise Stop('时间轴表内没有数据行，请先手动建立表结构。')
    last = data_rows[-1]

    # 模式一：闭合最后一条开放行（形如 "| 16:40- | … | 进行中 |"）
    if a.close:
        m = re.match(r'^\|\s*(\d{1,2}:\d{2})-\s*\|', lines[last])
        if not m:
            raise Stop(f'最后一条时间轴不是开放行（{lines[last][:24]}…），无需闭合。')
        start_mins = parse_hhmm(m.group(1), '开放行开始时刻')
        end_mins = parse_hhmm(a.close, '--close')
        if end_mins <= start_mins:
            raise Stop(f'闭合时刻 {a.close} 必须晚于开始 {m.group(1)}。')
        cells = [c.strip() for c in lines[last].strip('|').split('|')]
        cells[0] = f'{m.group(1)}-{a.close}'
        cells[1] = fmt_timeline_dur(end_mins - start_mins)
        if len(cells) >= 5:
            cells[4] = cells[4].replace('进行中', '').strip()
        lines[last] = '| ' + ' | '.join(cells) + ' |'
        print(f'→ 闭合时间轴：{cells[0]}（{cells[1]}）')
        if a.dry_run:
            print('（dry-run，未写入）')
            return
        save_diary(date, '\n'.join(lines))
        print('✅ 已写入日记（时间轴闭合）')
        return

    # 模式二：新增一行
    if not (a.start and a.cat and a.content):
        raise Stop('新增时间轴行需要 --start --cat --content（或 --close HH:MM 闭合开放行）。')
    s_mins = parse_hhmm(a.start, '--start')  # 无条件校验（开放行也要拦非法时刻）
    if a.end:
        e_mins = parse_hhmm(a.end, '--end')
        if e_mins <= s_mins:
            raise Stop(f'结束 {a.end} 必须晚于开始 {a.start}。')
        tcell, dur = f'{a.start}-{a.end}', fmt_timeline_dur(e_mins - s_mins)
    else:
        tcell, dur = f'{a.start}-', '—'
    mark = a.mark or ('进行中' if not a.end else '')
    row = f'| {tcell} | {dur} | {a.cat} | {a.content} | {mark} |'
    if re.match(r'^\|\s*\d{1,2}:\d{2}-\s*\|', lines[last]):
        open_cell = lines[last].split('|')[1].strip()
        print(f'⚠️ 上一条 "{open_cell}" 仍是开放行，本次未自动闭合；如需闭合先跑 --close。')
    print(f'→ 时间轴：{row}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    lines.insert(last + 1, row)
    save_diary(date, '\n'.join(lines))
    print('✅ 已写入日记（时间轴 +1 行）')


def simplify_train_content(content):
    """把详细训练内容（markdown表格/多行文本）简化为台账用的一行简版。"""
    if not content or str(content).strip() in ('', 'None', '待记录'):
        return ''
    text = str(content).strip()
    
    # 如果已经是简版（没有换行），直接返回
    if '\n' not in text:
        return text
    
    # 如果是 markdown 表格格式，提取动作名+重量+组次
    if '|' in text and '动作' in text and '重量' in text:
        lines = text.split('\n')
        actions = []
        for line in lines:
            line = line.strip()
            if not line.startswith('|') or '---' in line or '动作' in line:
                continue
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if len(cells) >= 3:
                name = cells[0]
                weight = cells[1]
                sets = cells[2]
                actions.append(f'{name}{weight}×{sets}')
        return '；'.join(actions)
    
    # 多行纯文本，合并成一行
    lines = [l.strip() for l in text.split('\n') if l.strip() and not l.startswith('#') and not l.startswith('|')]
    return '；'.join(lines)


def refresh_fitness_cost(wb):
    """全量重算「健身成本记录」sheet：从每日记录抽取训练数据，从财务明细累计健身支出。"""
    if '健身成本记录' not in wb.sheetnames:
        ws = wb.create_sheet('健身成本记录')
    else:
        ws = wb['健身成本记录']
    headers = SHEET_HEADERS['健身成本记录']
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)

    dws = wb['每日记录']
    daily = []
    for r in range(2, dws.max_row + 1):
        d = dws.cell(r, 1).value
        part = dws.cell(r, 9).value or ''
        cardio = dws.cell(r, 11).value
        if d and str(d).startswith('2026-'):
            daily.append((str(d)[:10], str(part).strip(), cardio))
    daily.sort(key=lambda x: x[0])
    if not daily:
        return

    start_date = daily[0][0]
    gym_expenses = {}
    fws = wb['财务明细']
    for r in range(2, fws.max_row + 1):
        d = fws.cell(r, 1).value
        n = fws.cell(r, 3).value or ''
        t = fws.cell(r, 4).value or ''
        a = fws.cell(r, 5).value or 0
        if d and str(d).startswith('2026') and '健身' in str(n) and str(t) == 'expense':
            dd = str(d)[:10]
            gym_expenses[dd] = gym_expenses.get(dd, 0) + float(a)

    # start_date 之前的健身支出作为初始累计（办卡发生在有每日记录之前）
    cum_expense = sum(v for d, v in gym_expenses.items() if d < start_date)
    cum_train = 0
    for row_idx, (date, part, cardio) in enumerate(daily, start=2):
        is_train = '是' if (part and part != '休息') else '否'
        if is_train == '是':
            cum_train += 1
        if date in gym_expenses:
            cum_expense += gym_expenses[date]
        total_visits = 54 + cum_train
        per_visit = round(cum_expense / total_visits, 2) if total_visits > 0 and cum_expense > 0 else 0
        ws.cell(row_idx, 1, date)
        ws.cell(row_idx, 2, is_train)
        ws.cell(row_idx, 3, part if part else '—')
        ws.cell(row_idx, 4, cardio if cardio is not None else 0)
        ws.cell(row_idx, 5, total_visits)
        ws.cell(row_idx, 6, round(cum_expense, 2))
        ws.cell(row_idx, 7, per_visit)
        ws.cell(row_idx, 8, f'健身支出 ¥{gym_expenses[date]:.2f}' if date in gym_expenses else None)


def cmd_train(a):
    date = a.date
    print(f'→ 训练：{a.part} | 有氧{a.cardio or 0}min')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    dws = wb['每日记录']
    r, _ = ensure_daily_row(dws, date)
    dws.cell(r, 9, a.part)
    if a.content:
        # 台账存简版，详细表格放日记
        simple = simplify_train_content(a.content)
        dws.cell(r, 10, simple)
    if a.cardio is not None:
        dws.cell(r, 11, a.cardio)
    if a.feeling:
        dws.cell(r, 12, a.feeling)
    # 自动换行
    for col in [10, 12]:
        dws.cell(r, col).alignment = Alignment(wrap_text=True, vertical='top')
    # 同步刷新健身成本记录
    refresh_fitness_cost(wb)
    # 找到当前日期在健身成本记录中的行，用于校验
    fc_row = None
    fcws = wb['健身成本记录']
    for rr in range(2, fcws.max_row + 1):
        if norm_date(fcws.cell(rr, 1).value) == date:
            fc_row = rr
            break
    checks = [('每日记录', r, 9, a.part, '训练部位')]
    if fc_row:
        checks.append(('健身成本记录', fc_row, 3, a.part if a.part else '—', '健身成本训练部位'))
    save_atomic(wb, checks)
    print('✅ 已同步：每日记录（训练列）+ 健身成本记录。日记的训练明细表格仍由我按你的描述详细撰写。')


def cmd_mantra(a):
    date = a.date
    note = a.note or '标准'
    print(f'→ 准提咒：{a.count}遍（{note}）')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    mws = wb['准提咒进度']
    exist = next((r for r in range(2, mws.max_row + 1) if norm_date(mws.cell(r, 1).value) == date), None)
    row = exist or mws.max_row + 1
    mws.cell(row, 1, date)
    mws.cell(row, 2, a.count)
    mws.cell(row, 3, note)
    sort_by_date(mws, desc=False)
    cum, cnt = 0, 0
    for r in range(2, mws.max_row + 1):
        d = mws.cell(r, 1).value
        if not d:
            continue
        cnt += 1
        v = mws.cell(r, 2).value
        cum += v if isinstance(v, (int, float)) else 0
        mws.cell(r, 4, cnt)
        mws.cell(r, 5, cum)
        mws.cell(r, 6, 10000)
        mws.cell(r, 7, f'{cum / 10000 * 100:.1f}%')
    dws = wb['每日记录']
    r2, _ = ensure_daily_row(dws, date)
    dws.cell(r2, 13, a.count)
    text = load_diary(date)
    text = diary_set_mantra(text, a.count, note, cum, 10000)
    save_atomic(wb, [('每日记录', r2, 13, a.count, '准提咒同步'),
                     ('准提咒进度', row if not exist else exist, 2, a.count, '遍数存在')])
    save_diary(date, text)
    print(f'✅ 已同步：准提咒进度(累计{cum}遍 {cum / 100:.1f}%) + 每日记录 + 日记')


def cmd_weight(a):
    date = a.date
    print(f'→ 体重：{a.kg}kg')
    body = None
    if a.body:
        body = [float(x) for x in a.body.split(',')]
        if len(body) != len(BODY_COLS):
            raise Stop(f'--body 需要{len(BODY_COLS)}个数字（逗号分隔，顺序：{",".join(BODY_COLS)}）')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    dws = wb['每日记录']
    r, _ = ensure_daily_row(dws, date)
    dws.cell(r, 3, a.kg)
    checks = [('每日记录', r, 3, a.kg, '体重同步')]
    extra = ''
    if body:
        bws = wb['体成分记录']
        br = next((i for i in range(2, bws.max_row + 1) if norm_date(bws.cell(i, 1).value) == date), None)
        if br is None:
            bws.insert_rows(2)
            bws.cell(2, 1, date)
            br = 2
        bws.cell(br, 2, a.kg)
        for i, v in enumerate(body, start=3):
            bws.cell(br, i, v)
        sort_by_date(bws, desc=True, ncol=len(SHEET_HEADERS['体成分记录']))
        br = next(i for i in range(2, bws.max_row + 1) if norm_date(bws.cell(i, 1).value) == date)
        checks.append(('体成分记录', br, 2, a.kg, '体成分同步'))
        extra = f'（体脂{body[1]}% 骨骼肌{body[4]}kg 内脏脂肪{body[9]}级）'
    text = load_diary(date)
    text = diary_add_weight(text, a.kg, extra)
    save_atomic(wb, checks)
    save_diary(date, text)
    print('✅ 已同步：每日记录体重 + 日记备注' + (' + 体成分记录全指标' if body else ''))
    # 7日滚动均值（2026-09-01 每日称重制配套）：单日±1kg是水分噪音，滚动均值看真速率
    try:
        from datetime import datetime, timedelta
        d0 = datetime.strptime(date, '%Y-%m-%d')
        pts = []
        for i in range(7):
            dd = (d0 - timedelta(days=i)).strftime('%Y-%m-%d')
            for rr in range(2, dws.max_row + 1):
                if norm_date(dws.cell(rr, 1).value) == dd:
                    v = dws.cell(rr, 3).value
                    if isinstance(v, (int, float)) and v > 0:
                        pts.append((dd, float(v)))
                    break
        if len(pts) >= 2:
            avg = sum(v for _, v in pts) / len(pts)
            span = f"{pts[-1][0][5:]}~{pts[0][0][5:]}"
            rate = ''
            if len(pts) >= 4:
                days = (datetime.strptime(pts[0][0], '%Y-%m-%d') - datetime.strptime(pts[-1][0], '%Y-%m-%d')).days or 1
                rate = f" | 区间速率 {(pts[0][1] - pts[-1][1]) / days * 7:+.2f}kg/周"
            print(f'📈 7日滚动均值：{avg:.2f}kg（{len(pts)}点，{span}{rate}）')
    except Exception:
        pass


def cmd_stock(a):
    wb = load()
    sws = wb['食物库存明细']
    row = match_stock(sws, a.food)
    if row is None:
        raise Stop(f'库存表中没有"{a.food}"。新食材请先告诉我，我加行后再扣。')
    name = sws.cell(row, 1).value
    cur = sws.cell(row, 3).value
    if not isinstance(cur, (int, float)):
        raise Stop(f'"{name}"当前库存是"{cur}"（非数字），请先告诉我实际数量（stock --set）。')
    if a.consume:
        after = round(cur - a.consume, 1)
        action = f'{name}：{after} ← {cur} - {a.consume}（{a.note or a.date}）'
    elif a.add:
        after = round(cur + a.add, 1)
        action = f'{name}：{after} ← {cur} + {a.add}（{a.note or "补货"}）'
    elif a.set is not None:
        after = a.set
        action = f'{name}：{after} ← {cur}（设为准确值）'
    else:
        raise Stop('请给 --consume / --add / --set 之一。')
    if after < 0:
        raise Stop(f'扣减后库存为负（{after}），说明记录有出入，先核对再操作。')
    # 计数型单位防呆：数量>20 疑似把克数填进"袋/个/盒"列（2026-08-24 帕斯雀300"袋"事故）
    unit = str(sws.cell(row, 4).value or '').strip()
    if any(unit.startswith(u) for u in ('个', '袋', '盒', '枚', '份', '根', '瓶', '条')) and not a.force:
        ops_val = a.consume or a.add or (a.set if a.set is not None else 0)
        if isinstance(ops_val, (int, float)) and abs(ops_val) > 20:
            raise Stop(f'⛔ "{name}"单位是"{unit}"（计数型），操作量 {ops_val} {unit} 明显偏大，'
                       f'疑似把克数填进了计数列（如 460g→460袋）。确认无误请加 --force。')
    print(f'→ 库存：{action}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    sws.cell(row, 3, after)
    sws.cell(row, 5, '⚠️ 已吃完' if after <= 0 else '✅ 充足')
    sws.cell(row, 6, a.date)
    text = load_diary(a.date)
    text = diary_add_stock(text, action)
    save_atomic(wb, [('食物库存明细', row, 3, after, '库存更新')])
    save_diary(a.date, text)
    print('✅ 已同步：食物库存明细 + 日记库存更新')


def cmd_food(a):
    nums = [float(x) for x in a.per100.split(',')]
    if len(nums) != 4:
        raise Stop('--per100 需要4个数字：热量,蛋白,脂肪,碳水（每100g），如 "240,28,13,2"')
    kcal, p, f_, c = nums
    wb = load()
    nws = wb['盒马营养成分']
    row = next((r for r in range(2, nws.max_row + 1) if str(nws.cell(r, 1).value or '').strip() == a.name), None)
    verb = '更新' if row else '新增'
    if row is None:
        row = nws.max_row + 1
    vals = [a.name, a.spec or '—', kcal, p, f_, c, a.unit or '100g',
            round(kcal * (parse_unit_grams(a.unit) or 100) / 100, 1),
            a.source or '实际包装标签', a.note or '']
    print(f'→ 营养成分{verb}：{a.name}（{kcal}/{p}/{f_}/{c} per100g）')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    for cidx, v in enumerate(vals, 1):
        nws.cell(row, cidx, v)
    save_atomic(wb, [('盒马营养成分', row, 1, a.name, '食物行存在'),
                     ('盒马营养成分', row, 3, kcal, '热量列正确')])
    print(f'✅ 已{verb}：{a.name}（严格10列格式，热量在第3列）')


# ==================== 努力度评估 ====================

def parse_sleep_h(text):
    """作息列 → 睡眠小时数。优先'合计/总约/='后的总时长，否则累加各段。兼容 '4h34min'。"""
    s = str(text or '')
    if not s or '待' in s:
        return None
    m = re.search(r'(?:合计约?|总约?)\s*(\d+)h\s*(\d{1,2})', s) or re.search(r'=\s*(\d+)h\s*(\d{1,2})', s)
    if m:
        return round(int(m.group(1)) + int(m.group(2)) / 60, 1)
    segs = re.findall(r'(\d+)h\s*(\d{1,2})m', s)
    if segs:
        return round(sum(int(a) + int(b) / 60 for a, b in segs), 1)
    return None


def anchor_score(x, anchors):
    """分段线性插值打分（anchors 升序，越界取端点值）。"""
    if x is None:
        return None
    if x <= anchors[0][0]:
        return anchors[0][1]
    if x >= anchors[-1][0]:
        return anchors[-1][1]
    for (x1, y1), (x2, y2) in zip(anchors, anchors[1:]):
        if x1 <= x <= x2:
            return round(y1 + (y2 - y1) * (x - x1) / (x2 - x1), 1)
    return None


def deeds_of(wb, date):
    """返回当日流水：[(类别, 领域, 性质, 内容, 时长)]"""
    ws = wb['成长流水']
    return [(str(ws.cell(r, 2).value or '').strip(), str(ws.cell(r, 3).value or '').strip(),
             str(ws.cell(r, 4).value or '').strip(), str(ws.cell(r, 5).value or ''),
             float(ws.cell(r, 6).value or 0))
            for r in range(2, ws.max_row + 1) if norm_date(ws.cell(r, 1).value) == date]


def grade_of(score):
    if score is None:
        return '—'
    for lo, name in GRADES:
        if score >= lo:
            return name
    return GRADES[-1][1]


def effort_calc(wb, date):
    """汇总五维 → 分数+完整度+虚度标记。铁律：未记录 ≠ 0分（防误伤），但完整度<80%不给等级。"""
    dws = wb['每日记录']
    r = next((rr for rr in range(2, dws.max_row + 1) if norm_date(dws.cell(rr, 1).value) == date), None)
    if r is None:
        raise Stop(f'{date} 在每日记录中无行，请先记录当日数据（meal/sleep 等）。')
    deeds = deeds_of(wb, date)
    deep = [d for d in deeds if d[0] == '深度成长']
    aff = [d for d in deeds if d[0] == '事务']
    data, dims = {}, {}
    # 2026-08-31 用户定规：深度成长未汇报 = 默认"无"（0h计分、正常评级），有做用户再补录
    data['深度成长h'] = round(sum(d[4] for d in deep), 2)
    data['事务h'] = round(sum(d[4] for d in aff), 2) if aff else None
    # 领域分布 + 输出占比（精力去向仪表盘）
    domains = {}
    for d in deep:
        key = d[1] or '未分领域'
        domains.setdefault(key, {'总': 0.0, '输出': 0.0})
        domains[key]['总'] += d[4]
        if d[2] == '输出':
            domains[key]['输出'] += d[4]
    data['领域'] = {k: (round(v['总'], 2), round(v['输出'], 2)) for k, v in domains.items()}
    out_h = sum(v['输出'] for v in domains.values())
    data['输出占比'] = round(out_h / data['深度成长h'], 2) if data['深度成长h'] else None
    dims['深度成长'] = anchor_score(data['深度成长h'], DEPTH_ANCHORS)
    train_raw = str(dws.cell(r, 9).value or '')
    if not train_raw or train_raw.startswith('待'):
        data['训练'], dims['训练'] = '未记录', None
    elif '跳训' in train_raw:
        # 2026-09-03 用户定规：跳训=计划训练未执行 → 0分+标记，区别于计划内休息
        data['训练'], dims['训练'] = '跳训', 0.0
    else:
        data['训练'] = '休息(计划内)' if '休息' in train_raw else '完成'
        dims['训练'] = 100.0  # 有记录即视为执行（强度细节看日记）
    mantra = dws.cell(r, 13).value
    if mantra in (None, '', '待'):
        data['准提咒'], dims['准提咒'] = '未做（附加项）', 0.0
    else:
        data['准提咒'] = int(mantra)
        dims['准提咒'] = anchor_score(float(mantra), [(0, 0), (54, 2.5), (108, 5.0)])
    data['睡眠h'] = parse_sleep_h(dws.cell(r, 2).value)
    dims['睡眠'] = anchor_score(data['睡眠h'], SLEEP_ANCHORS)
    kcal, prot = dws.cell(r, 5).value, dws.cell(r, 6).value
    if kcal and prot:
        t = TARGETS[PLAN]['rest' if data['训练'] in ('休息(计划内)', '跳训') else 'train']
        dev = abs(kcal - t['kcal']) / t['kcal']
        kcal_score = 50 if dev <= 0.15 else (25 if dev <= 0.30 else 0)
        lo, hi = t['protein']
        if lo <= prot <= hi:
            prot_score = 50
        elif prot < lo:
            prot_score = 35 if prot >= lo * 0.9 else (20 if prot >= lo * 0.75 else 0)
        else:
            prot_score = 25  # 超上限：过量摄入同样扣
        data['管理'] = f'{round(kcal, 0):.0f}kcal/P{prot}g'
        dims['自我管理'] = float(kcal_score + prot_score)
    else:
        data['管理'], dims['自我管理'] = '未记录', None
    got = {k: v for k, v in dims.items() if v is not None and k in EFFORT_WEIGHTS}
    completeness = round(sum(EFFORT_WEIGHTS[k] for k in got) / sum(EFFORT_WEIGHTS.values()), 2)
    base_score = round(sum(v * EFFORT_WEIGHTS[k] for k, v in got.items()), 1) if got else None
    bonus = round(sum(dims.get(k, 0) for k in EFFORT_BONUS), 1)
    score = round(base_score + bonus, 1) if base_score is not None else None
    depth_h = data['深度成长h'] or 0
    flags = []
    if data['训练'] == '跳训':
        flags.append('🔴 跳训：计划训练未执行（训练维度计0分）')
    if not deep:
        # 未汇报默认无：按 0h 计分，下面负面标记照常触发（用户可补录更正，2026-08-31 定规）
        flags.append('深度成长未汇报→按默认0h计分（有做可补录：deed --cat 深度成长）')
    elif depth_h < 1 and data['训练'] != '休息(计划内)':
        flags.append('🔴 虚度嫌疑：深度成长<1h且非休息日')
    if data['睡眠h'] is not None and data['睡眠h'] > 10 and depth_h < 1:
        flags.append('⚠️ 逃避模式嫌疑：睡眠>10h且深度成长<1h')
    if kcal and kcal > TARGETS[PLAN]['train']['kcal'] * 1.5 and depth_h < 1:
        flags.append('⚠️ 暴食补偿嫌疑：热量超标50%+且深度成长<1h')
    if depth_h >= 1 and data['输出占比'] is not None and data['输出占比'] < OUTPUT_RATIO_FLOOR:
        flags.append(f'⚠️ 只进不出：输出占比 {int(data["输出占比"] * 100)}% < {int(OUTPUT_RATIO_FLOOR * 100)}%（纯输入易虚胖，加输出型练习）')
    grade = grade_of(score) if completeness >= 0.8 else f'暂不评定（完整度{int(completeness * 100)}%）'
    return {'data': data, 'dims': dims, 'base_score': base_score, 'bonus': bonus, 'score': score, 'grade': grade,
            'completeness': completeness, 'flags': flags}


def diary_add_deed(text, cat, domain, io, content, hours):
    lines = text.split('\n')
    tag = f'｜{domain}' if domain else ''
    tag += f'（{io}）' if io else ''
    line = f'- {cat}{tag}｜{content}（{hours:g}h）'
    for i, l in enumerate(lines):
        if l.startswith('## 成长记录'):
            if i + 1 < len(lines) and '- 待记录' in lines[i + 1]:
                lines[i + 1] = line
            else:
                lines.insert(i + 1, line)
            return '\n'.join(lines)
    anchor = next((i for i, l in enumerate(lines) if l.startswith('## 饮食') or l.startswith('## 实际饮食记录')), None)
    if anchor is None:
        raise Stop('日记缺少"## 成长记录"与"## 饮食"锚点，请手动检查。')
    lines[anchor:anchor] = ['## 成长记录（今日做了什么）', line, '', '## 努力度评估', '- 待评估', '']
    return '\n'.join(lines)


def diary_set_effort(text, date, res):
    w = EFFORT_WEIGHTS
    lines_out = [f'- **{res["score"] if res["score"] is not None else "—"} 分 · {res["grade"]}**（基础 {res["base_score"]} + 附加 {res["bonus"]}，完整度 {int(res["completeness"] * 100)}%）',
                 '',
                 '| 维度 | 权重 | 数据 | 得分 |', '|------|------|------|------|']
    labels = {'深度成长': lambda d: f'{d["深度成长h"]}h' if d['深度成长h'] is not None else '未记录',
              '训练': lambda d: d['训练'],
              '睡眠': lambda d: f'{d["睡眠h"]}h' if d['睡眠h'] is not None else '未记录',
              '自我管理': lambda d: d['管理']}
    for k in ['深度成长', '训练', '睡眠', '自我管理']:
        v = res['dims'][k]
        lines_out.append(f'| {k} | {int(w[k] * 100)}% | {labels[k](res["data"])} | {v if v is not None else "缺" } |')
    lines_out.append(f'| 准提咒（附加） | +{EFFORT_BONUS["准提咒"]} | {res["data"]["准提咒"]}遍 | +{res["dims"]["准提咒"]} |')
    if res['data'].get('领域'):
        dom_str = ' / '.join(f'{k} {v[0]}h（出{v[1]}h）' for k, v in res['data']['领域'].items())
        ratio = res['data'].get('输出占比')
        ratio_str = f'，输出占比 {int(ratio * 100)}%' if ratio is not None else ''
        lines_out.append(f'| 精力去向 | — | {dom_str}{ratio_str} | — |')
    lines_out += [f'- {f}' for f in res['flags']]
    lines_out.append('')
    block = ['## 努力度评估'] + lines_out
    lines = text.split('\n')
    start = next((i for i, l in enumerate(lines) if l.startswith('## 努力度评估')), None)
    if start is not None:
        end = next((i for i in range(start + 1, len(lines)) if lines[i].startswith('## ')), len(lines))
        lines[start:end] = block
    else:
        anchor = next((i for i, l in enumerate(lines) if l.startswith('## 饮食') or l.startswith('## 实际饮食记录')), None)
        if anchor is None:
            raise Stop('日记缺少"## 努力度评估"锚点，请手动检查。')
        lines[anchor:anchor] = block
    return '\n'.join(lines)


def cmd_deed(a):
    if a.cat not in DEED_CATS:
        raise Stop(f'类别必须是：{"、".join(DEED_CATS)}')
    domain, io = a.domain or '', a.io or ''
    if a.cat == '深度成长':
        if not domain or not io:
            raise Stop('深度成长必须给 --domain（领域）和 --io（输入/输出）。领域清单：'
                       f'{"、".join(KNOWN_DOMAINS)}')
        if io not in IO_TYPES:
            raise Stop(f'--io 只能是：{"、".join(IO_TYPES)}')
        if domain not in KNOWN_DOMAINS and not a.allow_new_domain:
            raise Stop(f'领域"{domain}"不在清单中：{"、".join(KNOWN_DOMAINS)}。\n'
                       f'若是新领域，先问用户是否加入（确认后加 --allow-new-domain 并同步更新领域清单）。')
        if a.allow_new_domain and domain not in KNOWN_DOMAINS:
            print(f'🆕 新领域"{domain}"首次入账（记得更新规范文档领域清单）')
    display_io = f'｜{io}' if io else ''
    print(f'→ 成长流水：{a.date} | {a.cat}{domain and f"｜{domain}" or ""}{display_io} | {a.content}（{a.hours}h）')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    wb = load()
    dws = wb['成长流水']
    dws.insert_rows(2)
    for c, v in enumerate([a.date, a.cat, domain, io, a.content, round(a.hours, 2), a.output or '', a.note or ''], 1):
        dws.cell(2, c, v)
    sort_by_date(dws, desc=True)
    text = load_diary(a.date)
    text = diary_add_deed(text, a.cat, domain, io, a.content, a.hours)
    row = next(rr for rr in range(2, dws.max_row + 1)
               if norm_date(dws.cell(rr, 1).value) == a.date and dws.cell(rr, 5).value == a.content)
    save_atomic(wb, [('成长流水', row, 2, a.cat, '流水行存在')])
    save_diary(a.date, text)
    hint = '（深度成长已记，可跑 effort 评分）' if a.cat == '深度成长' else ''
    print(f'✅ 已同步：成长流水 + 日记成长记录{hint}')


def cmd_effort(a):
    wb = load()
    if a.week:
        from datetime import timedelta
        d0 = datetime.strptime(a.date, '%Y-%m-%d')
        monday = d0 - timedelta(days=d0.weekday())
        print(f'===== {monday.strftime("%Y-%m-%d")} 起一周努力度（只读） =====')
        depths, a_days, n = [], 0, 0
        for i in range(7):
            d = (monday + timedelta(days=i)).strftime('%Y-%m-%d')
            try:
                res = effort_calc(wb, d)
            except Stop:
                continue
            n += 1
            depth = res['data']['深度成长h']
            flag = next((f for f in res['flags'] if '虚度嫌疑' in f or '逃避' in f or '暴食' in f), '')
            print(f'  {d}  深度 {depth if depth is not None else "缺"}h | 综合 {res["score"] if res["score"] is not None else "—"} | {res["grade"]} {flag}')
            if depth is not None:
                depths.append(depth)
            if res['score'] is not None and res['score'] >= 80:
                a_days += 1
        if depths:
            print(f'均值 {sum(depths) / len(depths):.2f}h/天 | A级及以上（两成日）{a_days}/{n} 天')
        else:
            print('⚠️ 本周无任何深度成长记录，评估体系等待数据（deed --cat 深度成长）')
        return
    res = effort_calc(wb, a.date)
    print(f'===== {a.date} 努力度 =====')
    for k in ['深度成长', '训练', '睡眠', '自我管理']:
        v = res['dims'][k]
        label = {'深度成长': lambda d: f'{d["深度成长h"]}h' if d['深度成长h'] is not None else '未记录',
                 '训练': lambda d: d['训练'],
                 '睡眠': lambda d: f'{d["睡眠h"]}h' if d['睡眠h'] is not None else '未记录',
                 '自我管理': lambda d: d['管理']}[k](res['data'])
        print(f'  {k}（{int(EFFORT_WEIGHTS[k] * 100)}%） {label} → {v if v is not None else "缺"}')
    if res['dims'].get('准提咒', 0) > 0:
        print(f'  ✨ 准提咒（附加+{EFFORT_BONUS["准提咒"]}分） {res["data"]["准提咒"]}遍 → +{res["dims"]["准提咒"]}')
    else:
        print(f'  ✨ 准提咒（附加项） 今日未做 → +0（不扣分）')
    if res['data'].get('领域'):
        dom_str = ' / '.join(f'{k} {v[0]}h（出{v[1]}h）' for k, v in res['data']['领域'].items())
        ratio = res['data'].get('输出占比')
        ratio_str = f' → 输出占比 {int(ratio * 100)}%' if ratio is not None else ''
        print(f'  精力去向：{dom_str}{ratio_str}')
    print(f'  综合 {res["score"]} 分 · {res["grade"]} · 完整度 {int(res["completeness"] * 100)}%' if '暂不评定' in res['grade']
          else f'  综合 {res["score"]} 分 · {res["grade"]}（完整度 {int(res["completeness"] * 100)}%）')
    for f in res['flags']:
        print(f'  {f}')
    if a.dry_run:
        print('（dry-run，未写入）')
        return
    ews = wb['每日努力度']
    row = next((rr for rr in range(2, ews.max_row + 1) if norm_date(ews.cell(rr, 1).value) == a.date), None) \
        or ews.max_row + 1
    vals = [a.date, data_v(res['data'], '深度成长h'), data_v(res['data'], '事务h'), res['data']['训练'],
            res['data']['准提咒'], res['data']['睡眠h'], res['dims']['自我管理'], res['score'],
            f"{int(res['completeness'] * 100)}%", res['grade'], '；'.join(res['flags'])]
    for c, v in enumerate(vals, 1):
        ews.cell(row, c, v)
    sort_by_date(ews, desc=True)
    text = load_diary(a.date)
    text = diary_set_effort(text, a.date, res)
    row = next(rr for rr in range(2, ews.max_row + 1) if norm_date(ews.cell(rr, 1).value) == a.date)
    save_atomic(wb, [('每日努力度', row, 10, res['grade'], '等级存在')])
    save_diary(a.date, text)
    print(f'✅ 已同步：每日努力度 + 日记评估区')


def data_v(data, key):
    v = data.get(key)
    return v if v is not None else ''


# ==================== 校验与查看（只读） ====================

def check_date(wb, date, quiet=False):
    issues = []
    dws, fws, mws = wb['每日记录'], wb['财务明细'], wb['准提咒进度']
    try:
        r = find_daily_row(dws, date)
    except Stop:
        return [f'{date}：每日记录无此行']
    # 消费一致性
    sheet_sum = round(sum(float(fws.cell(fr, 5).value or 0) for fr in range(2, fws.max_row + 1)
                          if norm_date(fws.cell(fr, 1).value) == date
                          and str(fws.cell(fr, 4).value) == 'expense'), 2)
    daily_sum = round(float(dws.cell(r, 15).value or 0), 2)
    if abs(sheet_sum - daily_sum) > 0.01:
        issues.append(f'{date}：每日记录消费 ¥{daily_sum} ≠ 财务明细支出合计 ¥{sheet_sum}')
    # 日记营养 vs 每日记录
    p = diary_path(date)
    if os.path.exists(p):
        with open(p, encoding='utf-8') as fh:
            text = fh.read()
        totals = diary_totals(text)
        if totals is not None and date > LEGACY_UNTIL:
            kcal, pr, fa, ca = totals
            for label, got, col in [('热量', kcal, 5), ('蛋白', pr, 6), ('脂肪', fa, 7), ('碳水', ca, 8)]:
                sheet_v = round(float(dws.cell(r, col).value or 0), 1)
                if abs(sheet_v - round(got, 1)) > 0.2:
                    issues.append(f'{date}：{label} 日记{got} ≠ 台账{sheet_v}')
        total, diary_total = diary_expense_total(text)
        if diary_total is not None and abs(diary_total - sheet_sum) > 0.01:
            issues.append(f'{date}：日记消费合计 ¥{diary_total} ≠ 财务明细 ¥{sheet_sum}')
    # 作息/体重一致性（日记 ⇄ 每日记录）——2026-11-22 铁律：md 与 xlsx 必须逐项相等
    if os.path.exists(p) and date > LEGACY_UNTIL:
        lines = text.split('\n')
        s = next((i for i, l in enumerate(lines) if l.startswith('## 作息')), None)
        if s is not None:
            e = next((i for i, l in enumerate(lines) if i > s and l.startswith('## ')), len(lines))
            md_sleep = parse_sleep_hours('\n'.join(lines[s:e]))
            sheet_sleep = parse_sleep_hours(dws.cell(r, 2).value)
            if date >= SLEEP_CONSIST_FROM and md_sleep and sheet_sleep and abs(md_sleep - sheet_sleep) > 0.05:
                issues.append(f'{date}：睡眠 日记{md_sleep:.2f}h ≠ 台账{sheet_sleep:.2f}h')
        mw = re.search(r'称重[:：]\s*([0-9]+(?:\.[0-9]+)?)\s*kg', text)
        sheet_w = dws.cell(r, 3).value
        if mw:
            if not isinstance(sheet_w, (int, float)):
                issues.append(f'{date}：体重 日记{mw.group(1)}kg 但台账体重列为空')
            elif abs(float(mw.group(1)) - float(sheet_w)) > 0.001:
                issues.append(f'{date}：体重 日记{mw.group(1)}kg ≠ 台账{sheet_w}kg')
    # 准提咒
    m = next((mws.cell(mr, 2).value for mr in range(2, mws.max_row + 1)
              if norm_date(mws.cell(mr, 1).value) == date), None)
    d = dws.cell(r, 13).value
    if m is not None and str(d) not in (str(m), 'None', ''):
        if str(d) != str(m):
            issues.append(f'{date}：准提咒 每日记录{d} ≠ 进度表{m}')
    return issues


def parse_sleep_hours(text):
    """解析作息列文本，返回当日睡眠总小时数。

    支持格式：
    - '=7h21m'（合计） / '（共 13h57m）'（合计） / '(当日合计7h21m)'
    - '主睡眠 00:03→03:44(3h41m)'（段内标注，多段求和）
    - '主睡眠 6:00→10:45 + 晚睡眠 20:40→01:07'（时间段，跨午夜自动处理）
    """
    import re
    if not text:
        return 0.0
    t = str(text)
    # 1) 显式合计最高优先：当日总睡眠：**8h36m** / =7h21m / 共 13h57m / 当日合计7h21m
    m = re.search(r'(?:当日总睡眠|总睡眠|合计|共|=)[:：]?\s*\*{0,2}(\d+)\s*h\s*(\d+)?\s*m?', t)
    if m:
        return int(m.group(1)) + int(m.group(2) or 0) / 60
    # 2) 多段"约 XhXXm"求和（日记常态：主睡眠 + 补觉/午睡 分行写）
    segs2 = re.findall(r'约\s*(\d+)\s*h\s*(\d+)?\s*m', t)
    if segs2:
        return sum(int(h) + int(mm or 0) / 60 for h, mm in segs2)
    # 3) 段内标注：(3h41m) / (8h4min)，多段求和
    segs = re.findall(r'\((\d+)\s*h\s*(\d+)?\s*m', t)
    if segs:
        return sum(int(h) + int(mm or 0) / 60 for h, mm in segs)
    # 4) 时间段 HH:MM→HH:MM / HH:MM-HH:MM（可跨午夜）
    rngs = re.findall(r'(\d{1,2}):(\d{2})\s*[→\-]\s*(\d{1,2}):(\d{2})', t)
    total = 0.0
    for h1, m1, h2, m2 in rngs:
        s = int(h1) * 60 + int(m1)
        e = int(h2) * 60 + int(m2)
        if e <= s:
            e += 1440  # 跨午夜
        total += (e - s) / 60
    return total


def cmd_weekly(a):
    """生成周报告（周一至周日）。"""
    from datetime import datetime, timedelta
    wb = load()
    dws = wb['每日记录']
    bws = wb['体成分记录'] if '体成分记录' in wb.sheetnames else None
    dims_ws = wb['身材维度记录'] if '身材维度记录' in wb.sheetnames else None
    deed_ws = wb['成长流水']
    fws = wb['财务明细']

    # 确定周一和周日
    d = datetime.strptime(a.date, '%Y-%m-%d')
    monday = d - timedelta(days=d.weekday())
    sunday = monday + timedelta(days=6)
    mon_str = monday.strftime('%Y-%m-%d')
    sun_str = sunday.strftime('%Y-%m-%d')
    prev_monday = (monday - timedelta(days=7)).strftime('%Y-%m-%d')
    week_num = monday.isocalendar()[1]

    print(f'# 第 {week_num} 周周报（{mon_str} ~ {sun_str}）')
    print()

    # 收集每日数据
    rows = []
    for r in range(2, dws.max_row + 1):
        date_val = dws.cell(r, 1).value
        ds = str(date_val)[:10] if date_val else ''
        if mon_str <= ds <= sun_str:
            rows.append({
                'date': ds,
                'sleep': dws.cell(r, 2).value or '',
                'weight': dws.cell(r, 3).value,
                'kcal': float(dws.cell(r, 5).value or 0),
                'protein': float(dws.cell(r, 6).value or 0),
                'fat': float(dws.cell(r, 7).value or 0),
                'carbs': float(dws.cell(r, 8).value or 0),
                'train': dws.cell(r, 9).value or '',
                'cardio': int(dws.cell(r, 11).value or 0),
                'mantra': dws.cell(r, 13).value,
                'expense': float(dws.cell(r, 15).value or 0),
            })
    rows.sort(key=lambda x: x['date'])

    if not rows:
        print('本周无数据')
        return

    n_days = len(rows)

    # ===== 一、体重与体成分 =====
    print('## 一、体重与体成分')
    print()
    tw = None
    pw = None
    if bws is not None:
        # 找周初（本周内最早一次，周一有记录即周一）和周末后（周日之后~下周四内最早一次）的体成分
        # 周一晨起体重 vs 周末后首次晨起体重，反映这一周的真实变化
        # 2026-09-01 修正：收尾称重可能漂移到下周二/三（如 8/30 周的收尾称重落在 9/1），
        # 周末后窗口由"仅下周一"放宽到下周四；周初/周末后均取窗口内日期最早的一条
        start_body = None
        end_body = None
        end_limit = (monday + timedelta(days=10)).strftime('%Y-%m-%d')
        for r in range(2, bws.max_row + 1):
            ds = str(bws.cell(r, 1).value)[:10]
            if mon_str <= ds <= sun_str and (start_body is None or ds < start_body['date']):
                start_body = {
                    'date': ds,
                    'weight': bws.cell(r, 2).value,
                    'bmi': bws.cell(r, 3).value,
                    'body_fat': bws.cell(r, 4).value,
                    'muscle': bws.cell(r, 5).value,
                    'skeletal': bws.cell(r, 7).value,
                    'protein_rate': bws.cell(r, 11).value,
                    'visceral': bws.cell(r, 12).value,
                    'bmr': bws.cell(r, 15).value,
                    'body_age': bws.cell(r, 16).value,
                }
            if sun_str < ds <= end_limit and (end_body is None or ds < end_body['date']):
                end_body = {
                    'date': ds,
                    'weight': bws.cell(r, 2).value,
                    'bmi': bws.cell(r, 3).value,
                    'body_fat': bws.cell(r, 4).value,
                    'muscle': bws.cell(r, 5).value,
                    'skeletal': bws.cell(r, 7).value,
                    'protein_rate': bws.cell(r, 11).value,
                    'visceral': bws.cell(r, 12).value,
                    'bmr': bws.cell(r, 15).value,
                    'body_age': bws.cell(r, 16).value,
                }

        if start_body and end_body:
            pw = start_body  # 周初
            tw = end_body    # 周末后
            print(f'| 指标 | 周初（{pw["date"][5:]}） | 周末后（{tw["date"][5:]}） | 变化 |')
            print('|------|--------|--------|------|')
            def diff(cur, prev, unit=''):
                if cur is None or prev is None:
                    return '—'
                d = round(float(cur) - float(prev), 2)
                arrow = '⬆️' if d > 0 else ('⬇️' if d < 0 else '➡️')
                return f'{d:+g}{unit} {arrow}'

            print(f'| 体重 | {pw["weight"]}kg | {tw["weight"]}kg | {diff(tw["weight"], pw["weight"], "kg")} |')
            print(f'| 体脂率 | {pw["body_fat"]}% | {tw["body_fat"]}% | {diff(tw["body_fat"], pw["body_fat"], "%")} |')
            if tw['skeletal'] and pw['skeletal']:
                print(f'| 骨骼肌量 | {pw["skeletal"]}kg | {tw["skeletal"]}kg | {diff(tw["skeletal"], pw["skeletal"], "kg")} |')
            if tw['visceral'] and pw['visceral']:
                print(f'| 内脏脂肪 | {pw["visceral"]}级 | {tw["visceral"]}级 | 持平 |')
            if tw['protein_rate'] and pw['protein_rate']:
                print(f'| 蛋白质率 | {pw["protein_rate"]}% | {tw["protein_rate"]}% | {diff(tw["protein_rate"], pw["protein_rate"], "%")} |')
            if tw['body_age'] and pw['body_age']:
                print(f'| 身体年龄 | {pw["body_age"]}岁 | {tw["body_age"]}岁 | {diff(tw["body_age"], pw["body_age"], "岁")} |')
            if tw['bmr'] and pw['bmr']:
                print(f'| 基础代谢 | {pw["bmr"]}kcal | {tw["bmr"]}kcal | {diff(tw["bmr"], pw["bmr"], "")} |')
            print()

            # 分析
            weight_chg = float(tw['weight']) - float(pw['weight'])
            fat_chg = float(tw['body_fat']) - float(pw['body_fat']) if tw['body_fat'] and pw['body_fat'] else 0
            muscle_chg = float(tw['skeletal']) - float(pw['skeletal']) if tw['skeletal'] and pw['skeletal'] else 0

            if weight_chg < 0 and fat_chg > 0 and muscle_chg < 0:
                print('**结论：体重降了但质量不好。** 掉的重量以肌肉和水分为主，体脂率反而上升。根因通常是睡眠不足 + 蛋白没吃够，需要警惕。')
            elif weight_chg < 0 and fat_chg <= 0 and muscle_chg >= 0:
                print('**结论：理想减脂状态 ✅** 掉脂肪、保肌肉，这是最健康的减脂节奏。')
            elif weight_chg > 0 and muscle_chg > 0:
                print('**结论：增肌期正向发展** 体重上涨但以肌肉为主，力量训练有效。')
            print()

    # 每日晨重走势（2026-09-01 每日称重制配套）：体成分称重之外的日粒度体重补充
    # w_near 收集周一/周日各±3天窗口内的每日晨重，兜底锚点可用到周边界外的邻近数据
    mon_m3 = (monday - timedelta(days=3)).strftime('%Y-%m-%d')
    sun_p3 = (sunday + timedelta(days=3)).strftime('%Y-%m-%d')
    w_near = []
    for r in range(2, dws.max_row + 1):
        ds = str(dws.cell(r, 1).value or '')[:10]
        if not (mon_m3 <= ds <= sun_p3):
            continue
        try:
            v = float(dws.cell(r, 3).value)
        except (TypeError, ValueError):
            continue
        if v > 0:
            w_near.append((ds, v))
    w_near.sort()
    daily_w = [(ds, v) for ds, v in w_near if mon_str <= ds <= sun_str]
    if daily_w:
        trend = ' → '.join(f'{ds[5:]} {v:g}kg' for ds, v in daily_w)
        print(f'**每日晨重走势：** {trend}')
        w_avg = round(sum(v for _, v in daily_w) / len(daily_w), 2)
        print(f'**周均体重：** {w_avg}kg（本周 {len(daily_w)}/{n_days} 天有称重记录）')
        print()

    # 周内无配对体成分称重的兜底（2026-09-01 用户规则）：以临近周一/周日的每日晨重为主算周变化
    # 注：dict 需含下游用到的全部键（weight/skeletal/body_fat/body_age），缺失键会 KeyError
    if (pw is None or tw is None) and w_near:
        def _near(target):
            best = None
            for ds, v in w_near:
                dist = abs((datetime.strptime(ds, '%Y-%m-%d') - target).days)
                if best is None or dist < best[0]:
                    best = (dist, ds, v)
            return (best[1], best[2]) if best else None
        s_anchor = _near(monday)
        e_anchor = _near(sunday)
        if s_anchor and e_anchor and s_anchor[0] < e_anchor[0]:
            pw = {'date': s_anchor[0], 'weight': s_anchor[1], 'skeletal': None,
                  'body_fat': None, 'body_age': None}
            tw = {'date': e_anchor[0], 'weight': e_anchor[1], 'skeletal': None,
                  'body_fat': None, 'body_age': None}
            w_chg = round(tw['weight'] - pw['weight'], 2)
            w_arrow = '⬆️' if w_chg > 0 else ('⬇️' if w_chg < 0 else '➡️')
            print(f'| 指标 | 临近周一（{pw["date"][5:]}） | 临近周日（{tw["date"][5:]}） | 变化 |')
            print('|------|--------|--------|------|')
            print(f'| 体重（每日称重） | {pw["weight"]:g}kg | {tw["weight"]:g}kg | {w_chg:+g}kg {w_arrow} |')
            print()
            print(f'**说明：** 本周无配对的体成分称重，体重变化取每日晨重中临近周一（{pw["date"][5:]}）与临近周日（{tw["date"][5:]}）的记录对比。')
            print()

    # ===== 二、营养 =====
    print('## 二、营养')
    print()
    avg_kcal = round(sum(r['kcal'] for r in rows) / n_days, 1)
    avg_protein = round(sum(r['protein'] for r in rows) / n_days, 1)
    avg_fat = round(sum(r['fat'] for r in rows) / n_days, 1)
    avg_carbs = round(sum(r['carbs'] for r in rows) / n_days, 1)
    protein_ok = sum(1 for r in rows if 120 <= r['protein'] <= 140)
    fat_ok = sum(1 for r in rows if 55 <= r['fat'] <= 75)
    kcal_ok = sum(1 for r in rows if r['kcal'] <= 1980)  # 宽松达标（训练日1780+200余量，2026-09-23 缺口加大版）

    print('| 指标 | 周平均 | 目标（训练日/休息日） | 达标天数 |')
    print('|------|--------|----------------------|----------|')
    print(f'| 热量 | {avg_kcal} kcal | 1780 / 1530 | {kcal_ok}/{n_days} |')
    print(f'| 蛋白 | {avg_protein}g | 120-140g | {protein_ok}/{n_days} |')
    print(f'| 脂肪 | {avg_fat}g | 55-75g | {fat_ok}/{n_days} |')
    print(f'| 碳水 | {avg_carbs}g | 150 / 115 | — |')
    print()

    # 营养小结
    print('**📝 小结：**', end=' ')
    nutr_notes = []
    if protein_ok / n_days >= 0.8:
        nutr_notes.append('蛋白达标率不错，肌肉恢复有保障')
    elif protein_ok / n_days >= 0.5:
        nutr_notes.append('蛋白过半达标，还有提升空间')
    else:
        nutr_notes.append('蛋白达标率偏低，需要重点加强')

    if avg_fat > 75:
        nutr_notes.append('脂肪偏高，注意控油')
    elif avg_fat < 55:
        nutr_notes.append('脂肪偏低，激素合成可能受影响')
    else:
        nutr_notes.append('脂肪控制在合理区间')

    if avg_kcal > 2100:
        nutr_notes.append('热量整体超标，需要收紧')
    elif avg_kcal < 1500:
        nutr_notes.append('热量偏低，小心掉肌肉')
    else:
        nutr_notes.append('热量整体可控')

    print('；'.join(nutr_notes) + '。')
    print()

    # ===== 三、训练 =====
    print('## 三、训练')
    print()
    train_days = sum(1 for r in rows if r['train'] and '休息' not in r['train'])
    rest_days = sum(1 for r in rows if r['train'] and '休息' in r['train'])
    total_cardio = sum(r['cardio'] for r in rows)

    print('| 项目 | 数据 |')
    print('|------|------|')
    print(f'| 训练天数 | {train_days}/{n_days} 天 |')
    print(f'| 休息日 | {rest_days} 天 |')
    print(f'| 有氧累计 | {total_cardio} min（约 {round(total_cardio/60, 1)}h） |')

    # 主训部位统计
    parts = []
    for r in rows:
        if r['train'] and '休息' not in r['train']:
            parts.append(f'{r["date"][5:]} {r["train"]}')
    if parts:
        print(f'| 训练分布 | {", ".join(parts)} |')
    print()

    # 训练小结
    print('**📝 小结：**', end=' ')
    train_notes = []
    if train_days >= 6:
        train_notes.append('训练全勤，执行力很强')
    elif train_days >= 4:
        train_notes.append('训练频率合格')
    else:
        train_notes.append('训练出勤率偏低')

    main_train = sum(1 for r in rows if r['train'] and '主训' in r['train'])
    if main_train >= 3:
        train_notes.append(f'主训 {main_train} 次，强度够')
    elif main_train >= 1:
        train_notes.append(f'主训 {main_train} 次，强度中等')
    else:
        train_notes.append('主训偏少，以轻量为主')

    if total_cardio >= 200:
        train_notes.append('有氧量充足，心肺有收益')
    elif total_cardio >= 100:
        train_notes.append('有氧量适中')
    else:
        train_notes.append('有氧偏少，可以加一点')

    print('；'.join(train_notes) + '。')
    print()

    # ===== 四、睡眠 =====
    print('## 四、睡眠')
    print()
    import re
    sleep_hours_list = []
    good_sleep = 0
    for r in rows:
        total = parse_sleep_hours(r['sleep'])
        if total > 0:
            sleep_hours_list.append(total)
        elif '通宵' in r['sleep']:
            sleep_hours_list.append(0)

    if sleep_hours_list:
        avg_sleep = round(sum(sleep_hours_list) / len(sleep_hours_list), 1)
        good_sleep = sum(1 for h in sleep_hours_list if h >= 7)
        worst = min(sleep_hours_list)
        best = max(sleep_hours_list)
        print('| 指标 | 数据 |')
        print('|------|------|')
        print(f'| 日均睡眠 | {avg_sleep}h |')
        print(f'| ≥7h 天数 | {good_sleep}/{len(sleep_hours_list)} 天 |')
        print(f'| 最短 | {worst:g}h |')
        print(f'| 最长 | {best:g}h |')
    print()

    # 睡眠小结
    print('**📝 小结：**', end=' ')
    if sleep_hours_list:
        sleep_notes = []
        if good_sleep / len(sleep_hours_list) >= 0.7:
            sleep_notes.append('睡眠质量不错，恢复有保障')
        elif good_sleep / len(sleep_hours_list) >= 0.4:
            sleep_notes.append('睡眠时好时坏，需要更规律')
        else:
            sleep_notes.append('睡眠严重不足，是当前最大短板')

        if avg_sleep >= 7.5:
            sleep_notes.append('日均睡眠充足')
        elif avg_sleep >= 6:
            sleep_notes.append(f'日均 {avg_sleep}h，偏少但能撑')
        else:
            sleep_notes.append(f'日均仅 {avg_sleep}h，严重缺觉')

        if min(sleep_hours_list) < 5:
            sleep_notes.append('有通宵/极短睡眠日，波动太大')

        print('；'.join(sleep_notes) + '。')
    else:
        print('睡眠数据不足，建议每天记录。')
    print()

    # ===== 五、深度成长 =====
    print('## 五、深度成长')
    print()
    domains = {}
    total_in = 0
    total_out = 0
    for r in range(2, deed_ws.max_row + 1):
        ds = str(deed_ws.cell(r, 1).value)[:10]
        if mon_str <= ds <= sun_str:
            cat = deed_ws.cell(r, 2).value or ''
            if cat == '深度成长':
                domain = deed_ws.cell(r, 3).value or '未分类'
                io = deed_ws.cell(r, 4).value or ''
                hours = float(deed_ws.cell(r, 6).value or 0)
                if domain not in domains:
                    domains[domain] = {'输入': 0, '输出': 0}
                if io in ('输入', '输出'):
                    domains[domain][io] += hours
                    if io == '输入':
                        total_in += hours
                    else:
                        total_out += hours

    if domains:
        print('| 领域 | 输入(h) | 输出(h) | 合计 |')
        print('|------|---------|---------|------|')
        for dom, vals in sorted(domains.items(), key=lambda x: -(x[1]['输入'] + x[1]['输出'])):
            total = vals['输入'] + vals['输出']
            print(f'| {dom} | {vals["输入"]:g} | {vals["输出"]:g} | {total:g} |')
        grand_total = total_in + total_out
        out_ratio = int(total_out / grand_total * 100) if grand_total > 0 else 0
        print(f'| **合计** | **{total_in:g}** | **{total_out:g}** | **{grand_total:g}** |')
        print()
        print(f'输出占比：{out_ratio}%' + ('（建议 ≥20%，纯输入易虚胖）' if out_ratio < 20 else '✅'))
    print()

    # 深度成长小结
    print('**📝 小结：**', end=' ')
    if domains:
        growth_notes = []
        grand = total_in + total_out
        if grand >= 40:
            growth_notes.append(f'投入 {grand:g}h，强度很高')
        elif grand >= 20:
            growth_notes.append(f'投入 {grand:g}h，充实的一周')
        elif grand >= 10:
            growth_notes.append(f'投入 {grand:g}h，有在推进')
        else:
            growth_notes.append(f'仅 {grand:g}h，本周投入偏少')

        if out_ratio >= 30:
            growth_notes.append(f'输出占比 {out_ratio}%，学以致用做得好')
        elif out_ratio >= 15:
            growth_notes.append(f'输出占比 {out_ratio}%，还行但可以更多')
        else:
            growth_notes.append(f'输出占比仅 {out_ratio}%，光输入不输出容易"虚胖"')

        # 投入最多的领域
        top_dom = max(domains.keys(), key=lambda k: domains[k]['输入'] + domains[k]['输出'])
        growth_notes.append(f'主力领域：{top_dom}')

        print('；'.join(growth_notes) + '。')
    else:
        print('本周深度成长无记录。')
    print()

    # ===== 六、消费 =====
    print('## 六、消费')
    print()
    total_expense = round(sum(r['expense'] for r in rows), 2)
    print(f'- 本周总支出：**¥{total_expense}**')
    # 分类统计
    cats = {}
    for r in range(2, fws.max_row + 1):
        ds = str(fws.cell(r, 1).value)[:10]
        if mon_str <= ds <= sun_str:
            typ = str(fws.cell(r, 4).value or '')
            cat = fws.cell(r, 6).value or '其他'
            amt = float(fws.cell(r, 5).value or 0)
            if typ == 'expense' and amt > 0:  # 只算支出，不算退款和收入
                cats[cat] = cats.get(cat, 0) + amt
    if cats:
        cat_str = ' / '.join(f'{k} ¥{v:.0f}' for k, v in sorted(cats.items(), key=lambda x: -x[1]))
        print(f'- 分类：{cat_str}')
    print()

    # 消费小结
    print('**📝 小结：**', end=' ')
    exp_notes = []
    # 参考周预算 1500（粗略估算）
    weekly_budget = 1500
    if total_expense <= weekly_budget * 0.8:
        exp_notes.append('支出控制很好，远低于预算')
    elif total_expense <= weekly_budget:
        exp_notes.append('支出在预算内')
    elif total_expense <= weekly_budget * 1.2:
        exp_notes.append('支出略超预算，注意一下')
    else:
        exp_notes.append(f'超预算 {round(total_expense - weekly_budget)} 元，需要收紧')

    if cats:
        top_cat = max(cats.keys(), key=lambda k: cats[k])
        top_pct = int(cats[top_cat] / total_expense * 100)
        exp_notes.append(f'最大开销：{top_cat}（{top_pct}%）')

    print('；'.join(exp_notes) + '。')
    print()

    # ===== 七、压缩计划检查点 =====
    print('## 七、压缩计划检查点对照')
    print()
    print('| 检查点 | 目标 | 实际 | 状态 |')
    print('|--------|------|------|------|')

    # 周减重
    if tw and pw and tw['weight'] and pw['weight']:
        loss = round(float(pw['weight']) - float(tw['weight']), 2)
        if 0.6 <= loss <= 0.75:
            status = '✅'
        elif loss > 0.75:
            status = '⚠️ 太快（肌肉流失风险）'
        else:
            status = '⚠️ 未达标'
        print(f'| 周减重 | 0.6~0.75kg | {loss:+g}kg（正=减） | {status} |')
    else:
        print(f'| 周减重 | 0.6~0.75kg | 数据不足 | ⚠️ |')

    # 蛋白
    p_status = '✅' if protein_ok >= 6 else '❌'
    print(f'| 蛋白 ≥6 天达标 | ≥6/{n_days} | {protein_ok}/{n_days} | {p_status} |')

    # 睡眠
    if sleep_hours_list:
        s_status = '✅' if good_sleep >= 5 else '❌'
        print(f'| 睡眠 ≥7h ≥5 天 | ≥5/{len(sleep_hours_list)} | {good_sleep}/{len(sleep_hours_list)} | {s_status} |')
    else:
        print(f'| 睡眠 ≥7h ≥5 天 | ≥5/{n_days} | 数据不足 | ⚠️ |')

    # 骨骼肌量
    if tw and pw and tw['skeletal'] and pw['skeletal']:
        m_chg = round(float(tw['skeletal']) - float(pw['skeletal']), 2)
        m_status = '✅' if m_chg >= 0 else '❌'
        print(f'| 骨骼肌量不下降 | 持平/涨 | {m_chg:+g}kg | {m_status} |')
    else:
        print(f'| 骨骼肌量不下降 | 持平/涨 | 数据不足 | ⚠️ |')

    print()

    # 压缩计划小结
    print('**📝 小结：**', end=' ')
    check_pass = 0
    check_total = 0
    if tw and pw and tw['weight'] and pw['weight']:
        check_total += 1
        loss = round(float(pw['weight']) - float(tw['weight']), 2)
        if 0.6 <= loss <= 0.75:
            check_pass += 1
    check_total += 1
    if protein_ok >= 6:
        check_pass += 1
    if sleep_hours_list:
        check_total += 1
        if good_sleep >= 5:
            check_pass += 1
    if tw and pw and tw['skeletal'] and pw['skeletal']:
        check_total += 1
        if float(tw['skeletal']) >= float(pw['skeletal']):
            check_pass += 1

    if check_pass == check_total:
        print(f'{check_pass}/{check_total} 项全达标，完美执行压缩计划 ✅')
    elif check_pass >= check_total * 0.6:
        print(f'{check_pass}/{check_total} 项达标，整体可控，继续加油')
    else:
        print(f'仅 {check_pass}/{check_total} 项达标，执行力度不够，需要拉回正轨')
    print()

    print('---')
    print()

    # ===== 八、本周总结 =====
    print('## 八、本周总结')
    print()

    # --- 亮点 ---
    print('### ✨ 本周亮点')
    print()
    highlights = []

    # 体成分亮点
    if tw and pw:
        if tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) > float(pw['skeletal']):
            highlights.append(f'骨骼肌量上涨 {round(float(tw["skeletal"]) - float(pw["skeletal"]), 2)}kg，训练有效')
        if tw['body_fat'] and pw['body_fat'] and float(tw['body_fat']) < float(pw['body_fat']):
            highlights.append(f'体脂率下降 {round(float(pw["body_fat"]) - float(tw["body_fat"]), 1)}%，减脂见效')
        if tw['body_age'] and pw['body_age'] and float(tw['body_age']) < float(pw['body_age']):
            highlights.append(f'身体年龄年轻了 {int(float(pw["body_age"]) - float(tw["body_age"]))} 岁')

    # 训练亮点
    if train_days >= 5:
        highlights.append(f'训练 {train_days} 天，出勤率拉满')
    if total_cardio >= 200:
        highlights.append(f'有氧累计 {round(total_cardio/60, 1)}h，心肺耐力在线')

    # 营养亮点
    if protein_ok >= 5:
        highlights.append(f'蛋白达标 {protein_ok} 天，肌肉营养跟上了')
    if kcal_ok >= 6:
        highlights.append('热量控制稳定，几乎全周达标')

    # 深度成长亮点
    if domains:
        grand = total_in + total_out
        if grand >= 30:
            highlights.append(f'深度成长投入 {grand:g}h，相当充实的一周')
        out_ratio = int(total_out / grand * 100) if grand > 0 else 0
        if out_ratio >= 30:
            highlights.append(f'输出占比 {out_ratio}%，学以致用做得好')

    # 准提咒亮点
    mantra_days = sum(1 for r in rows if r['mantra'] and int(r['mantra'] or 0) >= 108)
    if mantra_days >= 5:
        highlights.append(f'准提咒 {mantra_days} 天满遍，功课稳定')

    if highlights:
        for i, h in enumerate(highlights, 1):
            print(f'{i}. {h}')
    else:
        print('本周数据较少，继续加油。')
    print()

    # --- 问题 ---
    print('### ⚠️ 本周问题')
    print()
    problems = []

    # 体成分问题
    if tw and pw and tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) < float(pw['skeletal']):
        problems.append(f'骨骼肌量下降 {round(float(pw["skeletal"]) - float(tw["skeletal"]), 2)}kg，有肌肉流失风险')
    if tw and pw and tw['body_fat'] and pw['body_fat'] and float(tw['body_fat']) > float(pw['body_fat']):
        problems.append(f'体脂率上升 {round(float(tw["body_fat"]) - float(pw["body_fat"]), 1)}%')

    # 营养问题
    if protein_ok < 5:
        problems.append(f'蛋白仅 {protein_ok}/{n_days} 天达标，肌肉恢复缺原料')
    if fat_ok < 4:
        problems.append(f'脂肪控制不佳，仅 {fat_ok}/{n_days} 天在目标区间')

    # 睡眠问题
    if sleep_hours_list and good_sleep < 5:
        problems.append(f'睡眠 ≥7h 仅 {good_sleep}/{len(sleep_hours_list)} 天，恢复不足')
    if sleep_hours_list and min(sleep_hours_list) < 5:
        problems.append(f'有睡眠不足 5h 的日子，严重影响恢复和代谢')

    # 训练问题
    if train_days <= 3:
        problems.append(f'训练仅 {train_days} 天，出勤率偏低')

    # 深度成长问题
    if domains:
        grand = total_in + total_out
        out_ratio = int(total_out / grand * 100) if grand > 0 else 0
        if out_ratio < 15:
            problems.append(f'输出占比仅 {out_ratio}%，光输入不输出容易"虚胖"')

    if problems:
        for i, p in enumerate(problems, 1):
            print(f'{i}. {p}')
    else:
        print('本周各项指标表现均衡，没有明显短板。')
    print()

    # --- 改进建议 ---
    print('### 🎯 下周改进建议')
    print()
    suggestions = []

    if protein_ok < 5:
        suggestions.append('**蛋白优先**：每餐保证 30-40g 蛋白，起床餐和训练后餐尤其重要，不够就补蛋白粉')
    if sleep_hours_list and good_sleep < 4:
        suggestions.append('**睡眠补课**：至少保证 4 天 ≥7h，睡前 1h 不碰手机，设个固定入睡时间')
    if fat_ok < 4:
        suggestions.append('**脂肪控制**：减少烹调用油和高脂零食，用鸡胸/鱼虾替代部分红肉')
    if tw and pw and tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) < float(pw['skeletal']):
        suggestions.append('**保住肌肉**：训练强度别掉、蛋白吃够、热量缺口别太大，优先保肌再减脂')
    if domains and total_in + total_out > 0 and int(total_out / (total_in + total_out) * 100) < 15:
        suggestions.append('**增加输出**：六爻研究可以尝试写卦例分析，剧本创作可以动笔写片段，输出才是真学到')
    if total_cardio < 100:
        suggestions.append('**加一点有氧**：每周至少 2 次有氧，每次 30min，爬坡或快走都行')

    if not suggestions:
        suggestions.append('继续保持当前节奏，稳扎稳打就是胜利。')

    for i, s in enumerate(suggestions, 1):
        print(f'{i}. {s}')
    print()

    print('---')


def cmd_monthly(a):
    """生成月报告（默认上个月：回顾 + 复盘 + 下月建议，不改台账数据）。"""
    import re
    from datetime import datetime, timedelta
    import calendar

    wb = load()
    dws = wb['每日记录']
    bws = wb['体成分记录'] if '体成分记录' in wb.sheetnames else None
    deed_ws = wb['成长流水']
    fws = wb['财务明细']

    # 月份：--month YYYY-MM，缺省 = 上个月
    if a.month:
        month_str = a.month
    else:
        today = datetime.now()
        month_str = (today.replace(day=1) - timedelta(days=1)).strftime('%Y-%m')
    year, mon = int(month_str[:4]), int(month_str[5:7])
    days_in_month = calendar.monthrange(year, mon)[1]
    first_str = f'{month_str}-01'
    last_str = f'{month_str}-{days_in_month:02d}'
    weeks = days_in_month / 7  # 周目标按此比例换算成月目标

    L = []
    L.append(f'# {month_str} 月报（{first_str} ~ {last_str}）')
    L.append('')

    # 收集每日数据
    rows = []
    for r in range(2, dws.max_row + 1):
        ds = str(dws.cell(r, 1).value or '')[:10]
        if first_str <= ds <= last_str:
            rows.append({
                'date': ds,
                'sleep': dws.cell(r, 2).value or '',
                'weight': dws.cell(r, 3).value,
                'kcal': float(dws.cell(r, 5).value or 0),
                'protein': float(dws.cell(r, 6).value or 0),
                'fat': float(dws.cell(r, 7).value or 0),
                'carbs': float(dws.cell(r, 8).value or 0),
                'train': dws.cell(r, 9).value or '',
                'cardio': int(dws.cell(r, 11).value or 0),
                'mantra': dws.cell(r, 13).value,
                'expense': float(dws.cell(r, 15).value or 0),
            })
    rows.sort(key=lambda x: x['date'])

    if not rows:
        print(f'{month_str} 无数据，月报未生成。')
        return
    n_days = len(rows)

    L.append(f'> 数据覆盖：{n_days}/{days_in_month} 天（完整度 {round(n_days / days_in_month * 100)}%）')
    L.append('')

    # ===== 一、体重与体成分 =====
    L.append('## 一、体重与体成分')
    L.append('')
    tw = pw = None
    if bws is not None:
        bodies = []
        for r in range(2, bws.max_row + 1):
            ds = str(bws.cell(r, 1).value or '')[:10]
            if first_str <= ds <= last_str:
                bodies.append({
                    'date': ds, 'weight': bws.cell(r, 2).value, 'bmi': bws.cell(r, 3).value,
                    'body_fat': bws.cell(r, 4).value, 'muscle': bws.cell(r, 5).value,
                    'skeletal': bws.cell(r, 7).value, 'protein_rate': bws.cell(r, 11).value,
                    'visceral': bws.cell(r, 12).value, 'bmr': bws.cell(r, 15).value,
                    'body_age': bws.cell(r, 16).value,
                })
        if bodies:
            bodies.sort(key=lambda x: x['date'])  # 体成分表倒序存储（最新在上），按日期升序排列
            pw, tw = bodies[0], bodies[-1]

            def fmt(v, unit=''):
                return f'{v}{unit}' if v is not None else '—'

            def diff(cur, prev, unit=''):
                if cur is None or prev is None:
                    return '—'
                d = round(float(cur) - float(prev), 2)
                arrow = '⬆️' if d > 0 else ('⬇️' if d < 0 else '➡️')
                return f'{d:+g}{unit} {arrow}'

            L.append(f'| 指标 | 月初（{pw["date"][5:]}） | 月末（{tw["date"][5:]}） | 变化 |')
            L.append('|------|--------|--------|------|')
            L.append(f'| 体重 | {fmt(pw["weight"], "kg")} | {fmt(tw["weight"], "kg")} | {diff(tw["weight"], pw["weight"], "kg")} |')
            L.append(f'| 体脂率 | {fmt(pw["body_fat"], "%")} | {fmt(tw["body_fat"], "%")} | {diff(tw["body_fat"], pw["body_fat"], "%")} |')
            if tw['skeletal'] and pw['skeletal']:
                L.append(f'| 骨骼肌量 | {pw["skeletal"]}kg | {tw["skeletal"]}kg | {diff(tw["skeletal"], pw["skeletal"], "kg")} |')
            if tw['visceral'] and pw['visceral']:
                L.append(f'| 内脏脂肪 | {pw["visceral"]}级 | {tw["visceral"]}级 | 持平 |')
            if tw['protein_rate'] and pw['protein_rate']:
                L.append(f'| 蛋白质率 | {pw["protein_rate"]}% | {tw["protein_rate"]}% | {diff(tw["protein_rate"], pw["protein_rate"], "%")} |')
            if tw['body_age'] and pw['body_age']:
                L.append(f'| 身体年龄 | {pw["body_age"]}岁 | {tw["body_age"]}岁 | {diff(tw["body_age"], pw["body_age"], "岁")} |')
            if tw['bmr'] and pw['bmr']:
                L.append(f'| 基础代谢 | {pw["bmr"]}kcal | {tw["bmr"]}kcal | {diff(tw["bmr"], pw["bmr"], "")} |')
            L.append(f'| 称重次数 | — | — | {len(bodies)} 次 |')
            L.append('')

            weight_chg = float(tw['weight']) - float(pw['weight'])
            has_fat = tw['body_fat'] is not None and pw['body_fat'] is not None
            has_muscle = tw['skeletal'] is not None and pw['skeletal'] is not None
            fat_chg = float(tw['body_fat']) - float(pw['body_fat']) if has_fat else 0
            muscle_chg = float(tw['skeletal']) - float(pw['skeletal']) if has_muscle else 0
            if weight_chg < 0 and has_fat and fat_chg <= 0 and has_muscle and muscle_chg >= 0:
                L.append('**结论：理想减脂状态 ✅** 掉脂肪、保肌肉，这是最健康的减脂节奏。')
            elif weight_chg < 0 and has_fat and fat_chg > 0 and has_muscle and muscle_chg < 0:
                L.append('**结论：减重质量不佳。** 掉的重量以肌肉和水分为主，体脂率反升，警惕（通常是睡眠不足+蛋白不够）。')
            elif weight_chg > 0 and has_muscle and muscle_chg > 0:
                L.append('**结论：增肌期正向发展**，体重上涨以肌肉为主。')
            else:
                parts = [f'体重 {weight_chg:+g}kg']
                if has_fat:
                    parts.append(f'体脂 {fat_chg:+g}%')
                if has_muscle:
                    parts.append(f'骨骼肌 {muscle_chg:+g}kg')
                if not has_fat and not has_muscle:
                    parts.append('（体脂/骨骼肌数据不全，仅看体重）')
                L.append('**结论：** ' + ' / '.join(parts) + '。')
            L.append('')
    if tw is None:
        L.append('本月无体成分记录（每周一晨起称重），数据不足。')
        L.append('')

    # ===== 二、营养 =====
    L.append('## 二、营养')
    L.append('')
    avg_kcal = round(sum(r['kcal'] for r in rows) / n_days, 1)
    avg_protein = round(sum(r['protein'] for r in rows) / n_days, 1)
    avg_fat = round(sum(r['fat'] for r in rows) / n_days, 1)
    avg_carbs = round(sum(r['carbs'] for r in rows) / n_days, 1)
    protein_ok = sum(1 for r in rows if 120 <= r['protein'] <= 140)
    fat_ok = sum(1 for r in rows if 55 <= r['fat'] <= 75)
    kcal_ok = sum(1 for r in rows if r['kcal'] <= 1980)  # 宽松达标（训练日1780+200余量，2026-09-23 缺口加大版）

    L.append('| 指标 | 月平均 | 目标（训练日/休息日） | 达标天数 |')
    L.append('|------|--------|----------------------|----------|')
    L.append(f'| 热量 | {avg_kcal} kcal | 1780 / 1530 | {kcal_ok}/{n_days} |')
    L.append(f'| 蛋白 | {avg_protein}g | 120-140g | {protein_ok}/{n_days} |')
    L.append(f'| 脂肪 | {avg_fat}g | 55-75g | {fat_ok}/{n_days} |')
    L.append(f'| 碳水 | {avg_carbs}g | 150 / 115 | — |')
    L.append('')

    nutr_notes = []
    if n_days and protein_ok / n_days >= 0.8:
        nutr_notes.append('蛋白达标率不错，肌肉恢复有保障')
    elif n_days and protein_ok / n_days >= 0.5:
        nutr_notes.append('蛋白过半达标，还有提升空间')
    else:
        nutr_notes.append('蛋白达标率偏低，需要重点加强')
    if avg_fat > 75:
        nutr_notes.append('脂肪偏高，注意控油')
    elif avg_fat < 55:
        nutr_notes.append('脂肪偏低，激素合成可能受影响')
    else:
        nutr_notes.append('脂肪控制在合理区间')
    if avg_kcal > 2100:
        nutr_notes.append('热量整体超标，需要收紧')
    elif avg_kcal < 1500:
        nutr_notes.append('热量偏低，小心掉肌肉')
    else:
        nutr_notes.append('热量整体可控')
    L.append('**📝 小结：** ' + '；'.join(nutr_notes) + '。')
    L.append('')

    # ===== 三、训练 =====
    L.append('## 三、训练')
    L.append('')
    train_days = sum(1 for r in rows if r['train'] and '休息' not in r['train'])
    rest_days = sum(1 for r in rows if r['train'] and '休息' in r['train'])
    total_cardio = sum(r['cardio'] for r in rows)
    part_count = {}
    for r in rows:
        if r['train'] and '休息' not in r['train']:
            part_count[r['train']] = part_count.get(r['train'], 0) + 1
    eff_weeks = n_days / 7  # 按实际有数据的天数折算目标（数据不全时不误判）
    train_target = max(1, round(4 * eff_weeks))

    L.append('| 项目 | 数据 |')
    L.append('|------|------|')
    L.append(f'| 训练天数 | {train_days}/{n_days} 天（月目标 ≥{train_target}） |')
    L.append(f'| 休息日 | {rest_days} 天 |')
    L.append(f'| 有氧累计 | {total_cardio} min（约 {round(total_cardio / 60, 1)}h） |')
    if part_count:
        dist = ' / '.join(f'{k}×{v}' for k, v in sorted(part_count.items(), key=lambda x: -x[1]))
        L.append(f'| 部位分布 | {dist} |')
    L.append('')

    train_notes = []
    if train_days >= train_target:
        train_notes.append('训练频率达标，出勤稳定')
    elif train_days >= round(3 * eff_weeks):
        train_notes.append('训练频率中等，略低于月目标')
    else:
        train_notes.append('训练出勤偏低，下月需要拉回')
    if total_cardio >= 800:
        train_notes.append('有氧量充足，心肺收益明显')
    elif total_cardio >= 400:
        train_notes.append('有氧量适中')
    else:
        train_notes.append('有氧偏少，建议每周至少 2 次 30min')
    L.append('**📝 小结：** ' + '；'.join(train_notes) + '。')
    L.append('')

    # ===== 四、睡眠 =====
    L.append('## 四、睡眠')
    L.append('')
    sleep_hours_list = []
    seg_days = 0
    for r in rows:
        total = parse_sleep_hours(r['sleep'])
        if total > 0:
            sleep_hours_list.append((r['date'], total))
        elif '通宵' in str(r['sleep']):
            sleep_hours_list.append((r['date'], 0))
        s = str(r['sleep'])
        if s and ('+' in s or ';' in s):
            seg_days += 1
    good_sleep = 0
    hours = []
    if sleep_hours_list:
        hours = [h for _, h in sleep_hours_list]
        avg_sleep = round(sum(hours) / len(hours), 1)
        good_sleep = sum(1 for h in hours if h >= 7)
        worst_d, worst = min(sleep_hours_list, key=lambda x: x[1])
        best_d, best = max(sleep_hours_list, key=lambda x: x[1])
        L.append('| 指标 | 数据 |')
        L.append('|------|------|')
        L.append(f'| 日均睡眠 | {avg_sleep}h |')
        L.append(f'| ≥7h 天数 | {good_sleep}/{len(hours)} 天 |')
        L.append(f'| 最短 | {round(worst, 1):g}h（{worst_d[5:]}） |')
        L.append(f'| 最长 | {round(best, 1):g}h（{best_d[5:]}） |')
        L.append(f'| 碎片化天数 | {seg_days} 天（多段睡眠） |')
    L.append('')

    if sleep_hours_list:
        sleep_notes = []
        if good_sleep / len(hours) >= 0.7:
            sleep_notes.append('睡眠整体达标，恢复有保障')
        elif good_sleep / len(hours) >= 0.4:
            sleep_notes.append('睡眠时好时坏，规律性不足')
        else:
            sleep_notes.append('睡眠严重不足，是本月最大短板')
        if avg_sleep >= 7.5:
            sleep_notes.append('日均睡眠充足')
        elif avg_sleep >= 6:
            sleep_notes.append(f'日均 {avg_sleep}h，偏少但能撑')
        else:
            sleep_notes.append(f'日均仅 {avg_sleep}h，长期缺觉拖垮训练和代谢')
        if seg_days >= len(hours) * 0.5:
            sleep_notes.append(f'碎片化睡眠 {seg_days} 天，建议固定起床时间调整生物钟')
        L.append('**📝 小结：** ' + '；'.join(sleep_notes) + '。')
    else:
        L.append('**📝 小结：** 睡眠数据不足，建议每天记录。')
    L.append('')

    # ===== 五、深度成长与修行 =====
    L.append('## 五、深度成长与修行')
    L.append('')
    domains = {}
    total_in = 0
    total_out = 0
    for r in range(2, deed_ws.max_row + 1):
        ds = str(deed_ws.cell(r, 1).value or '')[:10]
        if first_str <= ds <= last_str:
            cat = deed_ws.cell(r, 2).value or ''
            if cat == '深度成长':
                domain = deed_ws.cell(r, 3).value or '未分类'
                io = deed_ws.cell(r, 4).value or ''
                h = float(deed_ws.cell(r, 6).value or 0)
                if domain not in domains:
                    domains[domain] = {'输入': 0, '输出': 0}
                if io in ('输入', '输出'):
                    domains[domain][io] += h
                    if io == '输入':
                        total_in += h
                    else:
                        total_out += h

    if domains:
        L.append('| 领域 | 输入(h) | 输出(h) | 合计 |')
        L.append('|------|---------|---------|------|')
        for dom, vals in sorted(domains.items(), key=lambda x: -(x[1]['输入'] + x[1]['输出'])):
            tot = vals['输入'] + vals['输出']
            L.append(f'| {dom} | {vals["输入"]:g} | {vals["输出"]:g} | {tot:g} |')
        grand_total = total_in + total_out
        out_ratio = int(total_out / grand_total * 100) if grand_total > 0 else 0
        L.append(f'| **合计** | **{total_in:g}** | **{total_out:g}** | **{grand_total:g}** |')
        L.append('')
        L.append(f'输出占比：{out_ratio}%' + ('（建议 ≥20%，纯输入易虚胖）' if out_ratio < 20 else ' ✅'))
    else:
        grand_total = 0
        out_ratio = 0
        L.append('本月深度成长无记录。')
    L.append('')

    mantra_days = sum(1 for r in rows if r['mantra'] and int(r['mantra'] or 0) >= 108)
    mantra_total = sum(int(r['mantra'] or 0) for r in rows)
    L.append(f'修行：准提咒满遍 {mantra_days} 天，月累计 {mantra_total} 遍。')
    L.append('')

    growth_notes = []
    if domains:
        if grand_total >= round(20 * eff_weeks):
            growth_notes.append(f'投入 {grand_total:g}h，月度强度充实')
        elif grand_total >= round(10 * eff_weeks):
            growth_notes.append(f'投入 {grand_total:g}h，有在推进')
        else:
            growth_notes.append(f'仅 {grand_total:g}h，本月投入偏少')
        if out_ratio >= 30:
            growth_notes.append(f'输出占比 {out_ratio}%，学以致用做得好')
        elif out_ratio >= 15:
            growth_notes.append(f'输出占比 {out_ratio}%，还行但可以更多')
        else:
            growth_notes.append(f'输出占比仅 {out_ratio}%，光输入不输出容易"虚胖"')
        top_dom = max(domains.keys(), key=lambda k: domains[k]['输入'] + domains[k]['输出'])
        growth_notes.append(f'主力领域：{top_dom}')
    if mantra_days >= n_days * 0.5:
        growth_notes.append(f'准提咒满遍 {mantra_days} 天，功课稳定')
    elif mantra_days == 0:
        growth_notes.append('准提咒本月未坚持，下月设个固定时段')
    L.append('**📝 小结：** ' + '；'.join(growth_notes) + '。' if growth_notes else '**📝 小结：** 本月成长与修行数据较少。')
    L.append('')

    # ===== 六、消费（以财务明细为准） =====
    L.append('## 六、消费')
    L.append('')
    cats = {}
    total_income = 0
    for r in range(2, fws.max_row + 1):
        ds = str(fws.cell(r, 1).value or '')[:10]
        if first_str <= ds <= last_str:
            cat = str(fws.cell(r, 6).value or '其他')
            amt = float(fws.cell(r, 5).value or 0)
            if '收入' in cat:
                if amt > 0:
                    total_income += amt
                continue
            if amt > 0:
                cats[cat] = cats.get(cat, 0) + amt
            else:
                total_income += -amt  # 退款
    total_expense = round(sum(cats.values()), 2)
    monthly_budget = round(1500 * weeks)
    L.append(f'- 本月总支出：**¥{total_expense}**（月预算参考 ¥{monthly_budget}）')
    if total_income > 0:
        L.append(f'- 本月收入/退款：¥{round(total_income, 2)}')
    if cats:
        cat_str = ' / '.join(f'{k} ¥{v:.0f}' for k, v in sorted(cats.items(), key=lambda x: -x[1]))
        L.append(f'- 分类：{cat_str}')
    L.append('')

    exp_notes = []
    if total_expense <= monthly_budget * 0.8:
        exp_notes.append('支出控制很好，远低于预算')
    elif total_expense <= monthly_budget:
        exp_notes.append('支出在预算内')
    elif total_expense <= monthly_budget * 1.2:
        exp_notes.append('支出略超预算，注意一下')
    else:
        exp_notes.append(f'超预算 {round(total_expense - monthly_budget)} 元，需要收紧')
    if cats:
        top_cat = max(cats.keys(), key=lambda k: cats[k])
        top_pct = int(cats[top_cat] / total_expense * 100) if total_expense else 0
        exp_notes.append(f'最大开销：{top_cat}（{top_pct}%）')
    L.append('**📝 小结：** ' + '；'.join(exp_notes) + '。')
    L.append('')

    # ===== 七、月度检查点对照 =====
    L.append('## 七、月度检查点对照')
    L.append('')
    L.append('| 检查点 | 目标 | 实际 | 状态 |')
    L.append('|--------|------|------|------|')
    loss_target_lo = round(0.6 * eff_weeks, 1)
    loss_target_hi = round(0.75 * eff_weeks, 1)
    if loss_target_hi == 0:
        loss_target_hi = 0.1
    p_target = max(1, round(6 * eff_weeks))
    s_target = max(1, round(5 * eff_weeks))
    if tw and pw and tw['weight'] and pw['weight']:
        loss = round(float(pw['weight']) - float(tw['weight']), 2)
        if loss_target_lo <= loss <= loss_target_hi:
            st = '✅'
        elif loss > loss_target_hi:
            st = '⚠️ 太快（肌肉流失风险）'
        else:
            st = '⚠️ 未达标'
        L.append(f'| 月减重 | {loss_target_lo}~{loss_target_hi}kg | {loss:+g}kg（正=减） | {st} |')
    else:
        L.append(f'| 月减重 | {loss_target_lo}~{loss_target_hi}kg | 数据不足 | ⚠️ |')
    L.append(f'| 蛋白达标 | ≥{p_target}/{n_days} | {protein_ok}/{n_days} | {"✅" if protein_ok >= p_target else "❌"} |')
    if sleep_hours_list:
        L.append(f'| 睡眠 ≥7h | ≥{s_target} 天 | {good_sleep}/{len(hours)} | {"✅" if good_sleep >= s_target else "❌"} |')
    else:
        L.append(f'| 睡眠 ≥7h | ≥{s_target} 天 | 数据不足 | ⚠️ |')
    L.append(f'| 训练出勤 | ≥{train_target} 天 | {train_days}/{n_days} | {"✅" if train_days >= train_target else "❌"} |')
    if tw and pw and tw['skeletal'] and pw['skeletal']:
        m_chg = round(float(tw['skeletal']) - float(pw['skeletal']), 2)
        L.append(f'| 骨骼肌量不下降 | 持平/涨 | {m_chg:+g}kg | {"✅" if m_chg >= 0 else "❌"} |')
    else:
        L.append('| 骨骼肌量不下降 | 持平/涨 | 数据不足 | ⚠️ |')
    L.append(f'| 消费控制 | ≤¥{monthly_budget} | ¥{total_expense} | {"✅" if total_expense <= monthly_budget else "❌"} |')
    L.append('')

    check_pass = 0
    check_total = 0
    if tw and pw and tw['weight'] and pw['weight']:
        check_total += 1
        if loss_target_lo <= round(float(pw['weight']) - float(tw['weight']), 2) <= loss_target_hi:
            check_pass += 1
    check_total += 1
    if protein_ok >= p_target:
        check_pass += 1
    if sleep_hours_list:
        check_total += 1
        if good_sleep >= s_target:
            check_pass += 1
    check_total += 1
    if train_days >= train_target:
        check_pass += 1
    if tw and pw and tw['skeletal'] and pw['skeletal']:
        check_total += 1
        if float(tw['skeletal']) >= float(pw['skeletal']):
            check_pass += 1
    check_total += 1
    if total_expense <= monthly_budget:
        check_pass += 1

    if check_pass == check_total:
        L.append(f'**📝 小结：** {check_pass}/{check_total} 项全达标，完美执行月度计划 ✅')
    elif check_pass >= check_total * 0.6:
        L.append(f'**📝 小结：** {check_pass}/{check_total} 项达标，整体可控，继续加油')
    else:
        L.append(f'**📝 小结：** 仅 {check_pass}/{check_total} 项达标，执行力度不够，下月拉回正轨')
    L.append('')

    L.append('---')
    L.append('')

    # ===== 八、月度总结 =====
    L.append('## 八、月度总结')
    L.append('')
    L.append('### ✨ 本月亮点')
    L.append('')
    highlights = []
    if tw and pw:
        if tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) > float(pw['skeletal']):
            highlights.append(f'骨骼肌量上涨 {round(float(tw["skeletal"]) - float(pw["skeletal"]), 2)}kg，训练有效')
        if tw['body_fat'] and pw['body_fat'] and float(tw['body_fat']) < float(pw['body_fat']):
            highlights.append(f'体脂率下降 {round(float(pw["body_fat"]) - float(tw["body_fat"]), 1)}%，减脂见效')
        if tw['body_age'] and pw['body_age'] and float(tw['body_age']) < float(pw['body_age']):
            highlights.append(f'身体年龄年轻了 {int(float(pw["body_age"]) - float(tw["body_age"]))} 岁')
    if train_days >= train_target:
        highlights.append(f'训练 {train_days} 天，月度出勤达标')
    if total_cardio >= 800:
        highlights.append(f'有氧累计 {round(total_cardio / 60, 1)}h，心肺耐力在线')
    if protein_ok >= p_target:
        highlights.append(f'蛋白达标 {protein_ok} 天，肌肉营养跟上了')
    if kcal_ok >= max(1, round(6 * eff_weeks)):
        highlights.append('热量控制稳定，几乎全月达标')
    if domains and grand_total >= round(20 * eff_weeks):
        highlights.append(f'深度成长投入 {grand_total:g}h，充实的一个月')
    if domains and out_ratio >= 30:
        highlights.append(f'输出占比 {out_ratio}%，学以致用做得好')
    if mantra_days >= 10:
        highlights.append(f'准提咒 {mantra_days} 天满遍，功课稳定')

    if highlights:
        for i, h in enumerate(highlights, 1):
            L.append(f'{i}. {h}')
    else:
        L.append('本月数据较少，继续加油。')
    L.append('')

    L.append('### ⚠️ 本月问题')
    L.append('')
    problems = []
    if tw and pw and tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) < float(pw['skeletal']):
        problems.append(f'骨骼肌量下降 {round(float(pw["skeletal"]) - float(tw["skeletal"]), 2)}kg，有肌肉流失风险')
    if tw and pw and tw['body_fat'] and pw['body_fat'] and float(tw['body_fat']) > float(pw['body_fat']):
        problems.append(f'体脂率上升 {round(float(tw["body_fat"]) - float(pw["body_fat"]), 1)}%')
    if protein_ok < p_target * 0.7:
        problems.append(f'蛋白仅 {protein_ok}/{n_days} 天达标，肌肉恢复缺原料')
    if fat_ok < max(1, round(4 * eff_weeks)) * 0.7:
        problems.append(f'脂肪控制不佳，仅 {fat_ok}/{n_days} 天在目标区间')
    if sleep_hours_list and good_sleep < s_target:
        problems.append(f'睡眠 ≥7h 仅 {good_sleep}/{len(hours)} 天，恢复不足')
    if sleep_hours_list and min(hours) < 5:
        problems.append(f'有睡眠不足 5h 的日子（最短 {min(hours):g}h），严重影响恢复和代谢')
    if seg_days >= len(hours) * 0.5:
        problems.append(f'睡眠碎片化 {seg_days} 天，生物钟紊乱是根源')
    if train_days < round(3 * eff_weeks):
        problems.append(f'训练仅 {train_days} 天，出勤率偏低')
    if domains and out_ratio < 15:
        problems.append(f'输出占比仅 {out_ratio}%，光输入不输出容易"虚胖"')
    if total_expense > monthly_budget:
        problems.append(f'消费超预算 ¥{round(total_expense - monthly_budget)}')

    if problems:
        for i, p in enumerate(problems, 1):
            L.append(f'{i}. {p}')
    else:
        L.append('本月各项指标表现均衡，没有明显短板。')
    L.append('')

    L.append('### 🎯 下月改进建议')
    L.append('')
    suggestions = []
    if sleep_hours_list and good_sleep < s_target * 0.8:
        suggestions.append('**睡眠优先**：固定起床时间（比逼早睡有效 10 倍），先把日均拉到 7h，睡眠是当前一切指标的底层依赖')
    if seg_days >= len(hours) * 0.5:
        suggestions.append('**调整生物钟**：每天固定时间起床 + 起床后晒光，减少碎片化补觉')
    if protein_ok < p_target * 0.8:
        suggestions.append('**蛋白优先**：每餐保证 30-40g 蛋白，起床餐和训练后餐尤其重要，不够就补蛋白粉')
    if fat_ok < max(1, round(4 * eff_weeks)) * 0.8:
        suggestions.append('**脂肪控制**：减少烹调用油和高脂零食，用鸡胸/鱼虾替代部分红肉')
    if tw and pw and tw['skeletal'] and pw['skeletal'] and float(tw['skeletal']) < float(pw['skeletal']):
        suggestions.append('**保住肌肉**：训练强度别掉、蛋白吃够、热量缺口别太大，优先保肌再减脂')
    if train_days < round(3 * eff_weeks):
        suggestions.append('**训练出勤**：把训练固定到日程里（如周一/三/五+周末），降低启动摩擦')
    if domains and grand_total > 0 and out_ratio < 15:
        suggestions.append('**增加输出**：六爻研究写卦例分析，剧本创作动笔写片段，输出才是真学到')
    if total_cardio < 400:
        suggestions.append('**加一点有氧**：每周至少 2 次，每次 30min，爬坡或快走都行')
    if total_expense > monthly_budget:
        suggestions.append(f'**消费收紧**：下月控制在 ¥{monthly_budget} 内，先砍最大类目')

    if not suggestions:
        suggestions.append('继续保持当前节奏，稳扎稳打就是胜利。')

    for i, s in enumerate(suggestions, 1):
        L.append(f'{i}. {s}')
    L.append('')

    L.append('---')

    report = '\n'.join(L)
    print(report)

    # 保存文件
    import os
    os.makedirs(MONTHLY_REPORT_DIR, exist_ok=True)
    path = os.path.join(MONTHLY_REPORT_DIR, f'{month_str}月报.md')
    with open(path, 'w', encoding='utf-8') as fp:
        fp.write(report + '\n')
    print(f'\n✅ 月报已保存：{path}')


def cmd_check(a):
    wb = load()
    dates = [norm_date(wb['每日记录'].cell(r, 1).value) for r in range(2, wb['每日记录'].max_row + 1)]
    dates = [d for d in dates if d]
    targets = [a.date] if a.date else dates
    all_issues = []
    for d in targets:
        all_issues += check_date(wb, d)
    # 表头完整性：任何 sheet 的表头被顶掉/错位立即报警
    for name, expect in SHEET_HEADERS.items():
        ws = wb[name]
        got = [ws.cell(1, c).value for c in range(1, len(expect) + 1)]
        if got != expect:
            diff = [f'第{c}列 期望{e!r} 实得{g!r}' for c, (e, g) in enumerate(zip(expect, got), 1) if e != g]
            all_issues.append(f'{name}：表头异常（疑似整表错位/顶行）→ ' + '；'.join(diff))
    # 排序检查：每日记录/财务明细严格倒序（顶部最新），准提咒进度严格正序（累计才有意义）
    for name, desc in [('每日记录', True), ('财务明细', True), ('准提咒进度', False), ('体成分记录', True)]:
        ws = wb[name]
        ds = [norm_date(ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)]
        ds = [d for d in ds if re.match(r'^\d{4}-\d{2}-\d{2}', d)]
        if ds != sorted(ds, reverse=desc):
            direction = '倒序（顶部应最新）' if desc else '正序（旧→新，累计列才成立）'
            all_issues.append(f'{name}：日期不是严格{direction}')
    sws = wb['食物库存明细']
    for r in range(2, sws.max_row + 1):
        v = sws.cell(r, 3).value
        if isinstance(v, (int, float)) and v < 0:
            all_issues.append(f'库存为负：{sws.cell(r, 1).value} = {v}')
    if all_issues:
        print(f'❌ 发现 {len(all_issues)} 个不一致：')
        for i in all_issues:
            print(f'  - {i}')
        sys.exit(1)
    n_legacy = len([d for d in targets if d <= LEGACY_UNTIL])
    note = f'（{n_legacy} 个旧格式日期营养明细不全，以台账为准，未核对）' if n_legacy else ''
    print(f'✅ 校验通过：{len(targets)} 天数据一致，排序正确，库存无负数{note}')


def cmd_today(a):
    wb = load()
    dws, fws = wb['每日记录'], wb['财务明细']
    date = a.date
    r = find_daily_row(dws, date)
    print(f'===== {date} =====')
    print(f'作息：{dws.cell(r, 2).value}')
    print(f'体重：{dws.cell(r, 3).value or "—"}')
    print(f'饮食：{dws.cell(r, 4).value}')
    print(f'热量 {dws.cell(r, 5).value} / 蛋白 {dws.cell(r, 6).value} / 脂肪 {dws.cell(r, 7).value} / 碳水 {dws.cell(r, 8).value}')
    print(f'训练：{dws.cell(r, 9).value} | 有氧 {dws.cell(r, 11).value or 0}min')
    print(f'准提咒：{dws.cell(r, 13).value or "未记录"}遍')
    print(f'消费：{dws.cell(r, 14).value} = ¥{dws.cell(r, 15).value}')
    for fr in range(2, fws.max_row + 1):
        if norm_date(fws.cell(fr, 1).value) == date:
            print(f'  财务[{fws.cell(fr, 4).value}] {fws.cell(fr, 3).value} ¥{fws.cell(fr, 5).value}（{fws.cell(fr, 6).value}）')


def cmd_status(a):
    """输出常显信息（规范格式，直接贴回复末尾）。"""
    wb = load()
    dws = wb['每日记录']
    date = a.date
    r = find_daily_row(dws, date)

    kcal = float(dws.cell(r, 5).value or 0)
    protein = float(dws.cell(r, 6).value or 0)
    fat = float(dws.cell(r, 7).value or 0)
    carbs = float(dws.cell(r, 8).value or 0)
    train_part = dws.cell(r, 9).value or ''
    cardio = int(dws.cell(r, 11).value or 0)
    mantra = dws.cell(r, 13).value
    sleep_text = dws.cell(r, 2).value or ''

    # 目标直接读 TARGETS（2026-09-29 起改，杜绝规范更新后硬编码漂移复发；值：9/23 满配四支柱）
    # 只有明确标了"休息"才算休息日，空值/未记录按训练日算（计划6练1休）
    # ⚠️ 已知局限：跳训日台账训练列留空（9/23 裁定"跳训=休息"）→ 此处按训练日1780显示；
    #   跳训日正确目标为休息日1530，以日记营养汇总表为准（跳训低频，暂不解析日记判别）
    is_rest = '休息' in str(train_part) if train_part else False
    t_sel = TARGETS[PLAN]['rest' if is_rest else 'train']
    target_kcal = t_sel['kcal']
    target_carbs = t_sel['carb']
    # 脂肪区间（训练/休息同区间，2026-09-23 满配定值 55-75g 保激素）
    fat_low, fat_high = t_sel['fat']
    protein_low, protein_high = TARGETS[PLAN]['train']['protein']

    # 第一行：营养余额
    kcal_remain = round(target_kcal - kcal, 1)
    if kcal_remain >= 0:
        kcal_str = f'剩余 {kcal_remain:g} kcal'
    else:
        kcal_str = f'超 {abs(kcal_remain):g} kcal'

    if protein < protein_low:
        protein_str = f'蛋白 还需{protein_low - protein:g}g'
    elif protein > protein_high:
        protein_str = f'蛋白 超{protein - protein_high:g}g'
    else:
        protein_str = '蛋白 ✅'

    if fat < fat_low:
        fat_str = f'脂肪 还差{fat_low - fat:g}g'
    elif fat > fat_high:
        fat_str = f'脂肪 超{fat - fat_high:g}g'
    else:
        fat_str = '脂肪 ✅'

    carbs_remain = round(target_carbs - carbs, 1)
    if carbs_remain >= 0:
        carbs_str = f'碳水 剩{carbs_remain:g}g'
    else:
        carbs_str = f'碳水 超{abs(carbs_remain):g}g'

    line1 = f'{kcal_str} | {protein_str} | {fat_str} | {carbs_str}'

    # 第二行：管家提醒
    tips = []

    # 睡眠时长估算
    sleep_hours = None
    m = re.search(r'约\s*([\d.]+)\s*h', str(sleep_text))
    if m:
        sleep_hours = float(m.group(1))
    if sleep_hours is not None and sleep_hours < 6:
        tips.append(f'💤 睡眠不足（{sleep_hours:g}h），影响恢复')

    # 碳水不足提醒（训练日且碳水<一半目标）
    if not is_rest and carbs < target_carbs * 0.5:
        tips.append('⚡ 碳水偏低，训练日注意补碳')

    # 蛋白不足提醒
    if protein < protein_low:
        tips.append(f'🥩 蛋白还差 {protein_low - protein:g}g')

    # 准提咒提醒
    if not mantra or str(mantra) in ('0', '未记录', '待记录'):
        tips.append('🙏 准提咒还没完成')

    line2 = ' '.join(tips) if tips else ''

    # 第三行：进度激励
    progress = int(round(kcal / target_kcal * 100))
    if progress > 100:
        progress = 100
    # 根据训练状态动态生成激励
    if is_rest:
        motivate = '休息日好好恢复，肌肉在生长 🌱'
    elif cardio > 0:
        motivate = '今日已训练+有氧，脂肪在燃烧 🔥'
    elif train_part:
        motivate = '今日训练日，练完状态会更好 💪'
    else:
        motivate = '今日训练计划待定，加油完成 ✊'

    line3 = f'🎯 进度：{progress}% | {motivate}'

    # 第四行：健身成本核算（当年健身卡花费 / 当年训练次数）
    # 基准：台账建立前（2026-07-29 前）用户确认的历史次数 54 次（2026-08-29 口述存档）
    GYM_BASE_VISITS = 54
    year_prefix = date[:4]
    gym_cost = 0.0
    fws_stat = wb['财务明细']
    for rr in range(2, fws_stat.max_row + 1):
        ds_stat = str(fws_stat.cell(rr, 1).value or '')[:10]
        name_stat = str(fws_stat.cell(rr, 3).value or '')
        amt_stat = fws_stat.cell(rr, 5).value
        if ds_stat.startswith(year_prefix) and '健身' in name_stat:
            try:
                v = float(amt_stat or 0)
                if v > 0:
                    gym_cost += v
            except (TypeError, ValueError):
                pass
    gym_visits = GYM_BASE_VISITS
    for rr in range(2, dws.max_row + 1):
        ds_stat = str(dws.cell(rr, 1).value or '')[:10]
        tr_stat = str(dws.cell(rr, 9).value or '')
        if ds_stat.startswith(year_prefix) and tr_stat and '休息' not in tr_stat:
            gym_visits += 1
    if gym_visits > 0 and gym_cost > 0:
        per_visit = round(gym_cost / gym_visits, 1)
        line4 = f'💰 健身成本：¥{per_visit:g}/次（今年 {gym_visits} 次 · 共 ¥{round(gym_cost, 2):g}）'
    elif gym_cost > 0:
        line4 = f'💰 健身成本：今年已花 ¥{round(gym_cost, 2):g}，还没记录训练次数'
    else:
        line4 = ''

    # 输出
    print(f'--- 常显信息 · {date} ---')
    print(line1)
    if line2:
        print(f'💡 提醒：{line2}')
    print(line3)
    if line4:
        print(line4)


# ==================== CLI ====================

def main():
    ap = argparse.ArgumentParser(description='生活数据自动同步')
    sub = ap.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('meal')
    p.add_argument('--date', required=True)
    p.add_argument('--meal', required=True, help='起床餐/加餐/训练前/训练中/晚餐/夜宵...')
    p.add_argument('--food', required=True, help='营养成分表中的食物名')
    p.add_argument('--grams', type=float, help='克数（或ml）')
    p.add_argument('--count', type=float, help='个数/袋/盒/条，按常用分量换算')
    p.add_argument('--display', help='日记里的显示名，默认自动')
    p.set_defaults(fn=cmd_meal)

    p = sub.add_parser('expense')
    p.add_argument('--date', required=True)
    p.add_argument('--name', required=True)
    p.add_argument('--amount', type=float, required=True, help='正数支出；退款传负数')
    p.add_argument('--category', required=True)
    p.add_argument('--note', default='')
    p.set_defaults(fn=cmd_expense)

    p = sub.add_parser('income')
    p.add_argument('--date', required=True)
    p.add_argument('--name', required=True)
    p.add_argument('--amount', type=float, required=True)
    p.add_argument('--category', default=INCOME_DEFAULT_CAT)
    p.set_defaults(fn=cmd_income)

    p = sub.add_parser('sleep')
    p.add_argument('--date', required=True)
    p.add_argument('--range', required=True, help='如 8:27-16:31')
    p.add_argument('--label', help='主睡眠/午睡/傍晚睡，默认自动')
    p.set_defaults(fn=cmd_sleep)

    p = sub.add_parser('train')
    p.add_argument('--date', required=True)
    p.add_argument('--part', required=True, help='背+二头 / 肩 / 胸+三头 / 腿 / 休息')
    p.add_argument('--content', default='')
    p.add_argument('--cardio', type=int)
    p.add_argument('--feeling', default='')
    p.set_defaults(fn=cmd_train)

    p = sub.add_parser('timeline', help='时间轴补一行 / 闭合开放行（--close HH:MM）')
    p.add_argument('--date', required=True)
    p.add_argument('--start', help='开始 HH:MM，如 16:40')
    p.add_argument('--end', help='结束 HH:MM；缺省=开放行（进行中）')
    p.add_argument('--cat', help='类别：饮食/训练/恢复/事务/娱乐/修行/睡眠…')
    p.add_argument('--content', help='内容')
    p.add_argument('--mark', default='', help='标记：推定等')
    p.add_argument('--close', help='闭合最后一条开放行：只给结束时刻，如 16:50')
    p.set_defaults(fn=cmd_timeline)

    p = sub.add_parser('mantra')
    p.add_argument('--date', required=True)
    p.add_argument('--count', type=int, required=True)
    p.add_argument('--note', default='标准')
    p.set_defaults(fn=cmd_mantra)

    p = sub.add_parser('weight')
    p.add_argument('--date', required=True)
    p.add_argument('--kg', type=float, required=True)
    p.add_argument('--body', default='', help='体成分14项CSV：BMI,体脂率,肌肉量,肌肉率,骨骼肌量,骨骼肌率,骨量,水分率,蛋白质率,内脏脂肪,皮下脂肪率,去脂体重,基础代谢,身体年龄')
    p.set_defaults(fn=cmd_weight)
    p = sub.add_parser('dim')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'))
    p.add_argument('--chest', type=float, required=True, help='胸围cm')
    p.add_argument('--waist', type=float, required=True, help='腰围cm')
    p.add_argument('--hip', type=float, required=True, help='臀围cm')
    p.add_argument('--arm', type=float, required=True, help='手臂围cm')
    p.add_argument('--thigh', type=float, required=True, help='大腿围cm')
    p.add_argument('--calf', type=float, required=True, help='小腿围cm')
    p.add_argument('--note', default='')
    p.set_defaults(fn=cmd_dim)

    p = sub.add_parser('stock')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'))
    p.add_argument('--food', required=True)
    p.add_argument('--consume', type=float)
    p.add_argument('--add', type=float)
    p.add_argument('--set', type=float)
    p.add_argument('--note', default='')
    p.add_argument('--force', action='store_true', help='跳过计数型单位防呆检查（确认数量确实>20个/袋时用）')
    p.set_defaults(fn=cmd_stock)

    p = sub.add_parser('food')
    p.add_argument('--name', required=True)
    p.add_argument('--per100', required=True, help='"热量,蛋白,脂肪,碳水" 每100g')
    p.add_argument('--spec', default='')
    p.add_argument('--unit', default='100g', help='如 "40g/个" "50g(1包)"')
    p.add_argument('--source', default='实际包装标签')
    p.add_argument('--note', default='')
    p.set_defaults(fn=cmd_food)

    p = sub.add_parser('deed')
    p.add_argument('--date', required=True)
    p.add_argument('--cat', required=True, help='深度成长/事务/训练/修行/恢复/其他')
    p.add_argument('--domain', default='', help='领域资产池（深度成长必填）：剧本创作/系统建设/…')
    p.add_argument('--io', default='', help='输入/输出（深度成长必填）')
    p.add_argument('--allow-new-domain', action='store_true', help='新领域经用户确认后首次入账')
    p.add_argument('--content', required=True, help='做了什么，如"研究X并写报告"')
    p.add_argument('--hours', type=float, default=0.0)
    p.add_argument('--output', default='', help='产出/成果')
    p.add_argument('--note', default='')
    p.set_defaults(fn=cmd_deed)

    p = sub.add_parser('effort')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'))
    p.add_argument('--week', action='store_true', help='该日期所在周一至周日的汇总（只读）')
    p.set_defaults(fn=cmd_effort)

    p = sub.add_parser('weekly', help='生成周报告（周一至周日，只读）')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'), help='该日期所在周的周报')
    p.set_defaults(fn=cmd_weekly)

    p = sub.add_parser('monthly', help='生成月报告（默认上个月：回顾+复盘+下月建议，只读）')
    p.add_argument('--month', default=None, help='月份（YYYY-MM），默认上个月')
    p.set_defaults(fn=cmd_monthly)

    p = sub.add_parser('check')
    p.add_argument('--date', default=None)
    p.add_argument('--all', action='store_true')
    p.set_defaults(fn=cmd_check)

    p = sub.add_parser('today')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'))
    p.set_defaults(fn=cmd_today)

    p = sub.add_parser('status', help='常显信息（规范格式）')
    p.add_argument('--date', default=datetime.now().strftime('%Y-%m-%d'))
    p.set_defaults(fn=cmd_status)

    # 所有写命令统一支持 --dry-run
    for name, p in sub.choices.items():
        if name not in ('check', 'today', 'status', 'weekly', 'monthly'):
            p.add_argument('--dry-run', action='store_true', help='只打印动作，不写入')
            p.add_argument('--backfill', action='store_true', help='确认补录历史日期（绕过日期哨兵）')

    a = ap.parse_args()

    # ── 日期哨兵（2026-10-04 立，防"上下文冻结会话错写旧日期"复发）──
    # 背景：长期会话被压缩后续传时，摘要可能冻结在旧日期（如9/23），而真实已到10月初，
    # 多个会话曾把新数据错写进旧日期，git 史上修复过11+次。凡写入命令 --date 距真实今天
    # 超过3天且未显式加 --backfill，直接拒绝。物理闸门，不依赖任何会话"记得"。
    WRITE_CMDS = {'meal', 'expense', 'income', 'sleep', 'train', 'timeline',
                  'mantra', 'weight', 'dim', 'deed', 'effort'}
    if a.cmd in WRITE_CMDS and getattr(a, 'date', None):
        try:
            _pd = datetime.strptime(str(a.date)[:10], '%Y-%m-%d').date()
            _gap = (datetime.now().date() - _pd).days
        except ValueError:
            _gap = None
        if _gap is not None and _gap > 3 and not getattr(a, 'backfill', False):
            print(f'⛔ 日期哨兵：--date {a.date} 距真实今天 {datetime.now().date()} 已 {_gap} 天。')
            print('   这通常意味着当前会话的上下文日期过期了（不是真实日期）。')
            print('   → 先核实真实日期（date 命令 / ls 日记最新文件 / git log -1），')
            print('   → 如确认是补录历史数据，加 --backfill 重试。')
            sys.exit(3)
        if _gap is not None and _gap > 3:
            print(f'⚠️ 补录模式：--date {a.date}（{_gap} 天前），已确认补录。')

    try:
        a.fn(a)
    except Stop as e:
        print(f'⛔ {e}')
        sys.exit(2)


if __name__ == '__main__':
    main()
