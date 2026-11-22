# -*- coding: utf-8 -*-
"""
四柱八字排盘脚本（paipan.py）
依赖：lunar-python（已装机验证 1.4.8）
用途：把出生时间换算成完整命盘，输出 JSON。本脚本只产出「命盘事实」，不做任何吉凶解读。

用法：
  python paipan.py <年> <月> <日> <时> <分> <性别> [--jing 经度]

  - 性别：男=1 / 女=0，或直接写「男」/「女」
  - 时间：默认为北京时间(东八区，东经120°基准)。若需真太阳时，请显式传 --jing 出生地经度
    （东经为正，如北京 116.4；真太阳时 = 北京时间 + (经度-120)*4 分钟，脚本内部处理）

示例：
  python paipan.py 1990 5 20 14 30 男
"""
import json
import sys

from lunar_python import Solar

# ---------- 基础查表 ----------
GAN = "甲乙丙丁戊己庚辛壬癸"
ZHI = "子丑寅卯辰巳午未申酉戌亥"

GAN_WUXING = {g: w for g, w in zip(GAN, "木木火火土土金金水水")}
ZHI_WUXING = {z: w for z, w in zip(ZHI, "水土木木土火火土金金土水")}

# 天干五合
GAN_HE = {("甲", "己"): "土", ("乙", "庚"): "金", ("丙", "辛"): "水",
          ("丁", "壬"): "木", ("戊", "癸"): "火"}

# 地支六合
ZHI_HE = {("子", "丑"): "土", ("寅", "亥"): "木", ("卯", "戌"): "火",
          ("辰", "酉"): "金", ("巳", "申"): "水", ("午", "未"): "土"}

# 地支六冲
ZHI_CHONG = {("子", "午"), ("丑", "未"), ("寅", "申"), ("卯", "酉"),
             ("辰", "戌"), ("巳", "亥")}

# 地支六害（相穿）
ZHI_HAI = {("子", "未"), ("丑", "午"), ("寅", "巳"), ("卯", "辰"),
           ("申", "亥"), ("酉", "戌")}

# 地支相破
ZHI_PO = {("子", "酉"), ("卯", "午"), ("辰", "丑"), ("未", "戌"),
          ("寅", "亥"), ("巳", "申")}

# 三刑
XING_GROUPS = [("寅", "巳", "申"), ("丑", "戌", "未")]
ZIXING = set("辰午酉亥")

# 三合局
SANHE = {"水": ("申", "子", "辰"), "木": ("亥", "卯", "未"),
         "火": ("寅", "午", "戌"), "金": ("巳", "酉", "丑")}

# 十二长生（五行阴阳各表），表值按 长生/沐浴/冠带/临官/帝旺/衰/病/死/墓/绝/胎/养 顺序，
# 地支顺序为 亥子丑寅卯辰巳午未申酉戌（对应阳干），阴干逆行单独列。
CHANGSHENG_POS = ["长生", "沐浴", "冠带", "临官", "帝旺", "衰", "病", "死", "墓", "绝", "胎", "养"]
# 阳干（甲丙戊庚壬）与 阴支顺行起点
YANG_CS = {"甲": "亥", "丙": "寅", "戊": "寅", "庚": "巳", "壬": "申"}
YIN_CS = {"乙": "午", "丁": "酉", "己": "酉", "辛": "子", "癸": "卯"}
ZHI_ORDER = list(ZHI)      # 顺
ZHI_ORDER_REV = list(reversed(ZHI))  # 逆


def changsheng_of(gan, zhi):
    """返回某天干在某地支的十二长生状态。"""
    if gan in YANG_CS:
        start = ZHI_ORDER.index(YANG_CS[gan])
        seq = ZHI_ORDER[start:] + ZHI_ORDER[:start]
    else:
        start = ZHI_ORDER_REV.index(YIN_CS[gan])
        seq = ZHI_ORDER_REV[start:] + ZHI_ORDER_REV[:start]
    idx = seq.index(zhi)
    return CHANGSHENG_POS[idx]


def find_pair(a, b, table):
    if (a, b) in table:
        key = (a, b)
    elif (b, a) in table:
        key = (b, a)
    else:
        return None
    return table.get(key) if isinstance(table, dict) else key


def true_solar_time(hour, minute, jing):
    """按经度把北京时间校正为真太阳时（返回调整后的 (h,m)）。"""
    diff_min = round((jing - 120.0) * 4.0)
    total = hour * 60 + minute + diff_min
    total %= 24 * 60
    return total // 60, total % 60


def main():
    args = sys.argv[1:]
    jing = None
    liunian_year = None
    cleaned = []
    i = 0
    while i < len(args):
        if args[i] == "--jing" and i + 1 < len(args):
            jing = float(args[i + 1])
            i += 2
        elif args[i] == "--liunian" and i + 1 < len(args):
            liunian_year = int(args[i + 1])
            i += 2
        else:
            cleaned.append(args[i])
            i += 1

    if len(cleaned) < 6:
        print("用法: python paipan.py 年 月 日 时 分 性别 [--jing 经度]")
        sys.exit(1)

    y, mo, d, h, mi = map(int, cleaned[:5])
    g = cleaned[5]
    gender = 1 if g in ("男", "1", "male", "M") else 0

    if jing is not None:
        h, mi = true_solar_time(h, mi, jing)

    solar = Solar.fromYmdHms(y, mo, d, h, mi, 0)
    lunar = solar.getLunar()
    ec = lunar.getEightChar()

    # 四柱
    def hide_of(kind):
        m = {"year": ec.getYearHideGan, "month": ec.getMonthHideGan,
             "day": ec.getDayHideGan, "time": ec.getTimeHideGan}
        return m[kind]()

    def shishen_gan(kind):
        m = {"year": ec.getYearShiShenGan, "month": ec.getMonthShiShenGan,
             "day": ec.getDayShiShenGan, "time": ec.getTimeShiShenGan}
        return m[kind]()

    def shishen_zhi(kind):
        m = {"year": ec.getYearShiShenZhi, "month": ec.getMonthShiShenZhi,
             "day": ec.getDayShiShenZhi, "time": ec.getTimeShiShenZhi}
        return m[kind]()

    def naying(kind):
        m = {"year": ec.getYearNaYin, "month": ec.getMonthNaYin,
             "day": ec.getDayNaYin, "time": ec.getTimeNaYin}
        return m[kind]()

    def xunkong(kind):
        m = {"year": ec.getYearXunKong, "month": ec.getMonthXunKong,
             "day": ec.getDayXunKong, "time": ec.getTimeXunKong}
        return m[kind]()

    gan_map = {"year": ec.getYearGan, "month": ec.getMonthGan,
               "day": ec.getDayGan, "time": ec.getTimeGan}
    zhi_map = {"year": ec.getYearZhi, "month": ec.getMonthZhi,
               "day": ec.getDayZhi, "time": ec.getTimeZhi}

    pillars = {}
    for kind in ["year", "month", "day", "time"]:
        g_ = gan_map[kind]()
        z_ = zhi_map[kind]()
        pillars[kind] = {
            "gan": g_, "zhi": z_, "ganzhi": g_ + z_,
            "naying": naying(kind),
            "canggan": hide_of(kind),
            "shishen_gan": shishen_gan(kind),
            "shishen_zhi": shishen_zhi(kind),
            "xunkong": xunkong(kind),
        }

    # 五行统计（天干 + 地支本气 + 藏干加权：本气1 中气0.5 余气0.2）
    wuxing_tally = {"木": 0, "火": 0, "土": 0, "金": 0, "水": 0}
    for kind in ["year", "month", "day", "time"]:
        p = pillars[kind]
        wuxing_tally[GAN_WUXING[p["gan"]]] += 1
        wuxing_tally[ZHI_WUXING[p["zhi"]]] += 1
        cg = p["canggan"]
        for idx, c in enumerate(cg):
            w = 1.0 if idx == 0 else (0.5 if idx == 1 else 0.2)
            wuxing_tally[GAN_WUXING[c]] += w

    # 地支行冲合害刑破
    zhis = [pillars[k]["zhi"] for k in ["year", "month", "day", "time"]]
    zhi_rel = {"chong": [], "he": [], "hai": [], "po": [], "xing": [], "sanhe": []}
    pos = {"year": 0, "month": 1, "day": 2, "time": 3}
    for i1 in range(4):
        for i2 in range(i1 + 1, 4):
            a, b = zhis[i1], zhis[i2]
            r = find_pair(a, b, ZHI_CHONG)
            key = [k for k, v in pos.items() if v == i1][0] + "-" + [k for k, v in pos.items() if v == i2][0]
            if r is not None:
                zhi_rel["chong"].append(f"{key}({a}{b})")
            r = find_pair(a, b, ZHI_HE)
            if r is not None:
                zhi_rel["he"].append(f"{key}({a}{b}合{r})")
            r = find_pair(a, b, ZHI_HAI)
            if r is not None:
                zhi_rel["hai"].append(f"{key}({a}{b}穿)")
            r = find_pair(a, b, ZHI_PO)
            if r is not None:
                zhi_rel["po"].append(f"{key}({a}{b}破)")
    # 三刑（含自刑）
    seen_xing = set()
    for grp in XING_GROUPS:
        present = [z for z in grp if z in zhis]
        if len(present) == 3:
            seen_xing.add("三刑:" + "".join(present))
    if zhis[1] in ("子", "卯") and zhis[2] in ("子", "卯") and zhis[1] != zhis[2]:
        zhi_rel["xing"].append("子卯刑")
    for z in zhis:
        if z in ZIXING and zhis.count(z) >= 2:
            zhi_rel["xing"].append(f"{z}{z}自刑")
    zhi_rel["xing"].extend(sorted(seen_xing))
    # 三合 / 半合 / 暗拱（grp 中位即旺神，缺旺神两字相夹为暗拱）
    for wx, grp in SANHE.items():
        present = [z for z in grp if z in zhis]
        if len(present) == 3:
            zhi_rel["sanhe"].append(f"三合{wx}局({''.join(present)})")
        elif len(present) == 2:
            if grp[1] in present:
                zhi_rel["sanhe"].append(f"半合{wx}({''.join(present)})")
            else:
                zhi_rel["sanhe"].append(f"暗拱{grp[1]}({''.join(present)})")

    # 天干五合
    gans = [pillars[k]["gan"] for k in ["year", "month", "day", "time"]]
    gan_rel = []
    for i1 in range(4):
        for i2 in range(i1 + 1, 4):
            r = find_pair(gans[i1], gans[i2], GAN_HE)
            if r is not None:
                key = [k for k, v in pos.items() if v == i1][0] + "-" + [k for k, v in pos.items() if v == i2][0]
                gan_rel.append(f"{key}({gans[i1]}{gans[i2]}合{r})")

    # 十二长生（日主于四支）
    day_gan = pillars["day"]["gan"]
    cs = {k: changsheng_of(day_gan, pillars[k]["zhi"]) for k in ["year", "month", "day", "time"]}

    # 大运
    yun = ec.getYun(gender)
    dayun_list = []
    for d_yun in yun.getDaYun()[1:]:
        gz = d_yun.getGanZhi()
        age = d_yun.getStartAge()
        dayun_list.append({"ganzhi": gz, "age": age, "year": y + age - 1})

    out = {
        "公历": f"{y}-{mo:02d}-{d:02d} {h:02d}:{mi:02d}",
        "农历": lunar.toString(),
        "生肖": lunar.getYearShengXiao(),
        "性别": "男" if gender == 1 else "女",
        "真太阳时": "已校正" if jing is not None else "未校正(北京时间)",
        "四柱": pillars,
        "日主": day_gan,
        "日主五行": GAN_WUXING[day_gan],
        "日主十二长生": cs,
        "胎元": ec.getTaiYuan(), "胎元纳音": ec.getTaiYuanNaYin(),
        "命宫": ec.getMingGong(), "命宫纳音": ec.getMingGongNaYin(),
        "身宫": ec.getShenGong(), "身宫纳音": ec.getShenGongNaYin(),
        "五行统计": {k: round(v, 1) for k, v in wuxing_tally.items()},
        "地支关系": zhi_rel,
        "天干五合": gan_rel,
        "起运虚岁": dayun_list[0]["age"] if dayun_list else None,
        "大运": dayun_list,
    }

    if liunian_year is not None:
        liunian_info = None
        for d_yun in yun.getDaYun()[1:]:
            for ln in d_yun.getLiuNian():
                if ln.getYear() == liunian_year:
                    liunian_info = {
                        "年份": ln.getYear(),
                        "干支": ln.getGanZhi(),
                        "虚岁": ln.getAge(),
                        "所在大运": d_yun.getGanZhi(),
                    }
                    break
            if liunian_info:
                break
        out["流年"] = liunian_info

    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()