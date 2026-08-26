"""
MiMo Token Plan 用量分析工具

用法：
  python mimo-usage-checker.py                        # 自动获取最新官网数据 + 分析
  python mimo-usage-checker.py --login                # 弹出浏览器重新登录
  python mimo-usage-checker.py --xlsx <file.xlsx>     # 分析指定文件
"""
import asyncio
import json
import os
import re
import sys
import glob
import datetime
from pathlib import Path
from collections import defaultdict

CONFIG_DIR = os.path.expanduser("~/.mimo-usage-checker")
PLAN_INFO_FILE = os.path.join(CONFIG_DIR, "plan_info.json")
PHONE_FILE = os.path.join(CONFIG_DIR, "phone.txt")
COOKIE_FILE = os.path.join(CONFIG_DIR, "cookies.json")

# ===== 官方换算规则（不可改）=====
CREDIT_RATES = {
    "mimo-v2.5":     {"input_hit": 2,   "input_miss": 100, "output": 200},
    "mimo-v2.5-pro": {"input_hit": 2.5, "input_miss": 300, "output": 600},
    "mimo-v2.5-asr": {"audio_hour": 30_000_000},
}

# ===== 套餐额度（官方固定）=====
PLAN_QUOTAS = {
    "lite":     {"monthly": 4_100_000_000,   "yearly": 49_200_000_000},
    "standard": {"monthly": 11_000_000_000,  "yearly": 132_000_000_000},
    "pro":      {"monthly": 38_000_000_000,  "yearly": 456_000_000_000},
    "max":      {"monthly": 82_000_000_000,  "yearly": 984_000_000_000},
}

WEEKDAYS = ["一", "二", "三", "四", "五", "六", "日"]


def load_plan_info():
    if os.path.exists(PLAN_INFO_FILE):
        with open(PLAN_INFO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_plan_info(plan_name, quota, plan_type="monthly",
                   official_credits=None, valid_until=None):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    info = {"plan_name": plan_name, "quota": quota, "plan_type": plan_type}
    if official_credits is not None:
        info["official_credits"] = official_credits
    if valid_until is not None:
        info["valid_until"] = valid_until
    with open(PLAN_INFO_FILE, "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=2)


def get_plan(args):
    if "--plan" in args:
        i = args.index("--plan")
        if i + 1 < len(args):
            name = args[i + 1].lower()
            if name in PLAN_QUOTAS:
                return name, PLAN_QUOTAS[name]["monthly"]
            else:
                print(f"⚠️  未知套餐 '{name}'，可选: lite, standard, pro, max")
                sys.exit(1)
    info = load_plan_info()
    if info:
        name = info["plan_name"].lower()
        quota = info.get("quota")
        if name in PLAN_QUOTAS:
            return name, quota or PLAN_QUOTAS[name]["monthly"]
        if quota:
            return info["plan_name"], quota
    return None, None


def fmt(n):
    n = float(n)
    if n >= 1e9:  return f"{n/1e9:.2f}B"
    if n >= 1e6:  return f"{n/1e6:.1f}M"
    if n >= 1e3:  return f"{n/1e3:.1f}K"
    return f"{n:.0f}"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 核心：从 MiMo 控制台抓取最新官网数据（headless 模式）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async def fetch_official_data():
    """用已有 Cookie 在 headless 模式下抓取官网最新 Credits"""
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("❌ 需要安装: pip install playwright")
        return None

    if not os.path.exists(COOKIE_FILE):
        print("❌ 无已保存的 Cookie，请先运行: python mimo-usage-checker.py --login")
        return None

    with open(COOKIE_FILE, "r", encoding="utf-8") as f:
        cookies = json.load(f)

    result = {}
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context()
        await ctx.add_cookies(cookies)
        page = await ctx.new_page()
        await page.goto('https://platform.xiaomimimo.com/console/plan-manage')

        try:
            await page.wait_for_load_state("networkidle", timeout=20000)
            await asyncio.sleep(3)

            current_url = page.url
            if "/console/" not in current_url:
                print("⚠️  Cookie 已过期，请重新登录: python mimo-usage-checker.py --login")
                await browser.close()
                return None

            body_text = await page.inner_text('body')

            # 提取套餐名称
            plan_match = re.search(r'(Lite|Standard|Pro|Max)\s*(月度|年度)\s*套餐', body_text, re.IGNORECASE)
            if plan_match:
                result["plan_name"] = plan_match.group(1)
                result["plan_type"] = "monthly" if "月度" in plan_match.group(2) else "yearly"

            # 提取 "已使用 X%" 上方的 "used / total" 格式
            credits_match = re.search(
                r'([\d,]+)\s*/\s*([\d,]+)\s*\n\s*已使用\s*([\d.]+)%', body_text)
            if credits_match:
                result["official_credits"] = int(credits_match.group(1).replace(',', ''))
                result["quota"] = int(credits_match.group(2).replace(',', ''))
                result["used_pct"] = credits_match.group(3)

            # 提取有效期
            valid_match = re.search(r'有效期至\s*(\d{4}-\d{2}-\d{2})', body_text)
            if valid_match:
                result["valid_until"] = valid_match.group(1)

            # 提取 Token 总消耗
            token_match = re.search(r'Token 总消耗\s*\n\s*([\d,]+)\s*Tokens', body_text)
            if token_match:
                result["total_tokens"] = int(token_match.group(1).replace(',', ''))

        except Exception as e:
            print(f"⚠️  抓取失败: {e}")
        finally:
            await browser.close()

    return result if result else None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 分析 xlsx 并输出报告
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def analyze_xlsx(filepath, plan_override=None, quota_override=None,
                 official_credits=None, official_data=None):
    try:
        import openpyxl
    except ImportError:
        print("❌ 需要安装: pip install openpyxl")
        return

    # 从 official_data 提取信息
    if official_data:
        official_credits = official_data.get("official_credits", official_credits)
        if not plan_override:
            plan_override = official_data.get("plan_name")
        if not quota_override:
            quota_override = official_data.get("quota")

    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    # 按天+模型汇总
    daily = defaultdict(lambda: {
        "models": defaultdict(lambda: {
            "tokens": 0, "input_hit": 0, "input_miss": 0,
            "output": 0, "credits": 0, "requests": 0
        }),
        "total_tokens": 0, "input_hit": 0, "input_miss": 0,
        "output": 0, "credits": 0, "requests": 0
    })

    all_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        all_rows.append(row)
        date, model, total_tokens, input_hit, input_miss, output, audio_dur, req_count = row
        if model not in CREDIT_RATES:
            continue
        ih = int(input_hit or 0)
        im = int(input_miss or 0)
        ot = int(output or 0)
        tk = int(total_tokens or 0)
        rc = int(req_count or 0)
        rates = CREDIT_RATES[model]
        cr = ih * rates["input_hit"] + im * rates["input_miss"] + ot * rates["output"]

        d = str(date)
        daily[d]["total_tokens"] += tk
        daily[d]["input_hit"] += ih
        daily[d]["input_miss"] += im
        daily[d]["output"] += ot
        daily[d]["credits"] += cr
        daily[d]["requests"] += rc
        m = daily[d]["models"][model]
        m["tokens"] += tk; m["input_hit"] += ih; m["input_miss"] += im
        m["output"] += ot; m["credits"] += cr; m["requests"] += rc

    # 确定套餐
    if plan_override and quota_override:
        plan_name, quota = plan_override, quota_override
    elif plan_override:
        plan_name = plan_override
        quota = PLAN_QUOTAS.get(plan_override.lower(), {}).get("monthly", 0)
    else:
        plan_name, quota = get_plan(sys.argv[1:])

    if not plan_name:
        total_cr = sum(v["credits"] for v in daily.values())
        for name in ["lite", "standard", "pro", "max"]:
            if total_cr <= PLAN_QUOTAS[name]["monthly"]:
                plan_name = f"{name} (推测)"
                quota = PLAN_QUOTAS[name]["monthly"]
                break
        else:
            plan_name = "max (推测)"
            quota = PLAN_QUOTAS["max"]["monthly"]

    # xlsx 全量汇总
    total_tk = sum(v["total_tokens"] for v in daily.values())
    total_ih = sum(v["input_hit"] for v in daily.values())
    total_im = sum(v["input_miss"] for v in daily.values())
    total_ot = sum(v["output"] for v in daily.values())
    total_cr = sum(v["credits"] for v in daily.values())
    total_rc = sum(v["requests"] for v in daily.values())
    total_input = total_ih + total_im

    # 按订阅周期过滤（如果有有效期信息）
    valid_until_str = (official_data or {}).get("valid_until")
    cycle_start = None
    if valid_until_str:
        # 套餐有效期通常是 1 个月，倒推起始日
        valid_until = datetime.datetime.strptime(valid_until_str, "%Y-%m-%d").date()
        # 从页面提取的套餐名推断是月度还是年度
        if plan_type_yearly(plan_name, official_data):
            cycle_start = valid_until - datetime.timedelta(days=365)
        else:
            cycle_start = valid_until - datetime.timedelta(days=30)

    # 订阅周期内汇总
    cycle_tk = 0
    cycle_ih = 0
    cycle_im = 0
    cycle_ot = 0
    cycle_cr = 0
    cycle_rc = 0
    cycle_input = 0
    cycle_daily = {}

    for d_str, v in daily.items():
        in_cycle = True
        if cycle_start:
            try:
                d_date = datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
                in_cycle = d_date >= cycle_start
            except:
                pass
        if in_cycle:
            cycle_tk += v["total_tokens"]
            cycle_ih += v["input_hit"]
            cycle_im += v["input_miss"]
            cycle_ot += v["output"]
            cycle_cr += v["credits"]
            cycle_rc += v["requests"]
            cycle_input += v["input_hit"] + v["input_miss"]
            cycle_daily[d_str] = v

    # ===== 输出 =====
    print()
    print("=" * 90)
    print("  MiMo Token Plan 用量分析")
    print("=" * 90)
    if valid_until_str:
        print(f"  订阅周期: {cycle_start} ~ {valid_until_str}")
    print()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 核心分析一：换算一致性验证
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    print("━" * 90)
    print("  ① 换算一致性验证：xlsx 计算值 vs 官网实时显示值")
    print("━" * 90)
    print()

    print(f"  📊 xlsx 全量换算 Credits:              {total_cr:>15,.0f}")
    if cycle_start:
        print(f"  📊 订阅周期内换算 Credits:             {cycle_cr:>15,.0f}")
        print(f"     （过滤 {cycle_start} 之前的历史数据）")
    print()

    if official_credits is not None and official_credits > 0:
        print(f"  🌐 官网实时显示 Credits:               {official_credits:>15,.0f}")

        # 用订阅周期内的数据对比（更准确）
        compare_cr = cycle_cr if cycle_start else total_cr
        diff = compare_cr - official_credits
        diff_pct = abs(diff) / official_credits * 100 if official_credits else 0

        if cycle_start:
            print(f"  📐 对比基准（周期内 xlsx）:             {compare_cr:>15,.0f}")
        print(f"  📏 差异（xlsx - 官网）:                  {diff:>+15,.0f}  ({diff_pct:+.2f}%)")
        print()

        if abs(diff) < official_credits * 0.01:
            print("  ✅ 一致：xlsx 换算结果与官网显示基本一致（差异 < 1%）")
            print("     结论：换算公式正确，数据无异常。")
        elif diff > 0:
            print("  ⚠️  xlsx 计算值 > 官网显示值")
            print("     可能原因：")
            print("     1. 非高峰时段（00:00-08:00）0.8x 系数未在 xlsx 中体现")
            print("     2. 官网可能排除了某些调试/测试请求")
            print("     3. xlsx 导出与官网统计存在时间差（~5分钟延迟）")
        else:
            print("  ⚠️  xlsx 计算值 < 官网显示值")
            print("     可能原因：")
            print("     1. xlsx 导出延迟，未包含最近的消耗")
            print("     2. 官网可能包含了 xlsx 未记录的其他消耗来源")
            print("     3. 换算规则可能有更新")
    else:
        print("  ❌ 未能获取官网 Credits 数据")
        print("     请检查 Cookie 是否有效，或重新登录: --login")
    print()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 核心分析二：每日消耗分析（仅订阅周期内）
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    print("━" * 90)
    print("  ② 每日消耗分析" + ("（订阅周期内）" if cycle_start else ""))
    print("━" * 90)
    print()

    display_daily = cycle_daily if cycle_start else daily
    print(f"  {'日期':<12} {'周':>2} {'Tokens':>12} {'缓存命中':>8} {'未命中':>10} {'输出':>8} {'Credits':>14} {'请求':>6}")
    print(f"  {'─'*12} {'─'*2} {'─'*12} {'─'*8} {'─'*10} {'─'*8} {'─'*14} {'─'*6}")

    for d in sorted(display_daily.keys()):
        v = display_daily[d]
        dt = datetime.datetime.strptime(d, "%Y-%m-%d")
        wd = WEEKDAYS[dt.weekday()]
        ti = v["input_hit"] + v["input_miss"]
        hit_pct = f"{v['input_hit']/ti*100:.0f}%" if ti else "—"
        print(f"  {d:<12} {wd:>2} {v['total_tokens']:>12,} {hit_pct:>8} {v['input_miss']:>10,} {v['output']:>8,} {v['credits']:>14,.0f} {v['requests']:>6}")

    # 汇总行
    d_tk = sum(v["total_tokens"] for v in display_daily.values())
    d_ih = sum(v["input_hit"] for v in display_daily.values())
    d_im = sum(v["input_miss"] for v in display_daily.values())
    d_ot = sum(v["output"] for v in display_daily.values())
    d_cr = sum(v["credits"] for v in display_daily.values())
    d_rc = sum(v["requests"] for v in display_daily.values())
    d_inp = d_ih + d_im
    days = len(display_daily)

    print(f"  {'─'*12} {'─'*2} {'─'*12} {'─'*8} {'─'*10} {'─'*8} {'─'*14} {'─'*6}")
    hit_pct_all = f"{d_ih/d_inp*100:.0f}%" if d_inp else "—"
    print(f"  {'合计':<12} {'':>2} {d_tk:>12,} {hit_pct_all:>8} {d_im:>10,} {d_ot:>8,} {d_cr:>14,.0f} {d_rc:>6}")
    if days > 0:
        print(f"  {'日均':<12} {'':>2} {d_tk//days:>12,} {'':>8} {'':>10} {'':>8} {d_cr/days:>14,.0f} {d_rc//days:>6}")
    print()

    # 趋势
    if days >= 2:
        daily_list = sorted(display_daily.items())
        credits_values = [v["credits"] for _, v in daily_list]
        avg_cr = sum(credits_values) / len(credits_values)
        max_day = max(daily_list, key=lambda x: x[1]["credits"])
        min_day = min(daily_list, key=lambda x: x[1]["credits"])

        print("  📈 趋势分析:")
        print(f"     日均 Credits: {avg_cr:,.0f}")
        print(f"     最高日: {max_day[0]} ({max_day[1]['credits']:,.0f} Credits, {max_day[1]['requests']:,} 请求)")
        print(f"     最低日: {min_day[0]} ({min_day[1]['credits']:,.0f} Credits, {min_day[1]['requests']:,} 请求)")
        if min_day[1]["credits"] > 0:
            print(f"     波动范围: {max_day[1]['credits']/min_day[1]['credits']:.1f}x")
    print()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 补充：按模型汇总
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    print("━" * 90)
    print("  ③ 按模型汇总")
    print("━" * 90)
    print()
    print(f"  {'模型':<18} {'Tokens':>14} {'命中':>12} {'未命中':>10} {'输出':>10} {'Credits':>14} {'占比':>6} {'请求':>6}")
    print(f"  {'─'*18} {'─'*14} {'─'*12} {'─'*10} {'─'*10} {'─'*14} {'─'*6} {'─'*6}")

    model_agg = defaultdict(lambda: {
        "tokens": 0, "input_hit": 0, "input_miss": 0,
        "output": 0, "credits": 0, "requests": 0
    })
    for v in display_daily.values():
        for model, md in v["models"].items():
            for k in model_agg[model]:
                model_agg[model][k] += md[k]

    for model, s in sorted(model_agg.items(), key=lambda x: -x[1]["credits"]):
        p = s["credits"] / d_cr * 100 if d_cr else 0
        print(f"  {model:<18} {s['tokens']:>14,} {s['input_hit']:>12,} {s['input_miss']:>10,} {s['output']:>10,} {s['credits']:>14,.0f} {p:>5.1f}% {s['requests']:>6}")

    print()
    print("  换算规则:")
    print("    mimo-v2.5:     命中×2    未命中×100   输出×200")
    print("    mimo-v2.5-pro: 命中×2.5  未命中×300   输出×600")
    print("    非高峰(00:00-08:00 北京) 0.8x 系数")
    print("=" * 90)
    print()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # ④ 套餐用量预估（基于官网实时数据 + 订阅周期）
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    print("━" * 90)
    print("  ④ 套餐用量预估" + ("（基于官网实时数据）" if official_credits else ""))
    print("━" * 90)
    print()

    today = datetime.date.today()

    if valid_until_str and official_credits is not None:
        # 基于官网实时数据计算
        valid_until = datetime.datetime.strptime(valid_until_str, "%Y-%m-%d").date()
        days_total = (valid_until - cycle_start).days if cycle_start else 30
        days_used = (today - cycle_start).days if cycle_start else today.day
        days_remaining = (valid_until - today).days

        daily_burn = official_credits / days_used if days_used > 0 else 0
        projected = official_credits / days_used * days_total if days_used > 0 else 0
        remaining = quota - official_credits

        print(f"  订阅周期:   {cycle_start} ~ {valid_until_str}（共 {days_total} 天）")
        print(f"  已过天数:   第 {days_used}/{days_total} 天（剩余 {days_remaining} 天）")
        print(f"  官网已用:   {fmt(official_credits)} Credits ({official_credits/quota*100:.1f}%)")
        print(f"  剩余额度:   {fmt(remaining)} Credits")
        print(f"  日均消耗:   {fmt(daily_burn)} Credits/天")
        print(f"  预估周期末: {fmt(projected)} Credits ({projected/quota*100:.1f}%)")

        if days_remaining > 0:
            days_left = remaining / daily_burn if daily_burn > 0 else float('inf')
            if days_left < days_remaining:
                print(f"  ⚠️  按当前速率，约 {days_left:.0f} 天后用完（还剩 {days_remaining} 天）")
            else:
                print(f"  ✅ 按当前速率，周期内够用（预计剩余 {days_left - days_remaining:.0f} 天余量）")
    elif days > 0:
        # 降级：用 xlsx 数据估算
        days_passed = today.day
        daily_burn = d_cr / days_passed if days_passed > 0 else 0
        remaining = quota - d_cr
        days_in_month = (today.replace(month=today.month % 12 + 1, day=1) - datetime.timedelta(days=1)).day if today.month < 12 else 31
        days_rem = days_in_month - today.day
        projected = d_cr / days_passed * days_in_month if days_passed > 0 else 0

        print(f"  当前天数:  月度第 {days_passed}/{days_in_month} 天")
        print(f"  已用:      {fmt(d_cr)} Credits ({d_cr/quota*100:.1f}%)")
        print(f"  剩余:      {fmt(remaining)} Credits")
        print(f"  日均消耗:  {fmt(daily_burn)} Credits/天")
        print(f"  预估月末:  {fmt(projected)} Credits ({projected/quota*100:.1f}%)")
        if daily_burn > 0:
            days_left = remaining / daily_burn
            if days_left < days_rem:
                print(f"  ⚠️  按当前速率，约 {days_left:.0f} 天后用完")
            else:
                print(f"  ✅ 按当前速率，月底前够用")
    print()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 引导用户继续提问
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    print("━" * 90)
    print("  💬 还需要分析什么数据？可以直接问我，例如：")
    print("━" * 90)
    print()
    print("  • 「对比上个周期的消耗变化」")
    print("  • 「哪天消耗最高？分析原因」")
    print("  • 「缓存命中率对 Credits 的影响有多大？」")
    print("  • 「按小时分析消耗分布」")
    print("  • 「如果换成 Pro 套餐要花多少钱？」")
    print("  • 「帮我写一个消耗预警规则」")
    print("  • 「导出一份 Markdown 报告」")
    print()


def plan_type_yearly(plan_name, official_data):
    """从 official_data 判断是否年度套餐"""
    if official_data and official_data.get("plan_type") == "yearly":
        return True
    return False


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 登录流程
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async def login():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("❌ pip install playwright && playwright install chromium")
        sys.exit(1)

    os.makedirs(CONFIG_DIR, exist_ok=True)
    dl = str(Path.home() / "Downloads")

    print("🚀 浏览器即将弹出，请用手机号验证登录...\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        ctx = await browser.new_context(accept_downloads=True)

        # 加载已保存的 Cookie
        if os.path.exists(COOKIE_FILE):
            with open(COOKIE_FILE, "r", encoding="utf-8") as f:
                cookies = json.load(f)
            await ctx.add_cookies(cookies)
            print("  🍪 已加载保存的 Cookie\n")

        page = await ctx.new_page()
        await page.goto('https://platform.xiaomimimo.com/console/plan-manage')

        try:
            # 等待页面加载，判断是否需要登录
            await page.wait_for_load_state("networkidle", timeout=15000)
            await asyncio.sleep(2)

            current_url = page.url
            if "/console/" not in current_url:
                print("🔐 需要登录，正在切换到手机号验证...\n")
                sms_selectors = [
                    'text=手机号登录',
                    'text=短信登录',
                    'text=短信验证登录',
                    'text=短信验证码登录',
                    '[class*="tab"]:has-text("手机")',
                    '[class*="tab"]:has-text("短信")',
                    'div[role="tab"]:has-text("手机")',
                    'span:has-text("手机号")',
                ]
                clicked = False
                for sel in sms_selectors:
                    try:
                        loc = page.locator(sel).first
                        if await loc.is_visible(timeout=2000):
                            await loc.click()
                            print(f"  ✅ 已切换到手机号验证\n")
                            clicked = True
                            break
                    except Exception:
                        continue
                if not clicked:
                    print("  ⚠️  未找到手机号登录选项，请手动选择手机号验证方式\n")

                # 自动填入已保存的手机号
                if os.path.exists(PHONE_FILE):
                    with open(PHONE_FILE, "r", encoding="utf-8") as f:
                        saved_phone = f.read().strip()
                    if saved_phone:
                        phone_inputs = page.locator('input[placeholder*="手机"], input[placeholder*="电话"], input[type="tel"], input[name*="phone"], input[name*="mobile"]')
                        if await phone_inputs.count() > 0:
                            await phone_inputs.first.fill(saved_phone)
                            print(f"  📱 已自动填入手机号: {saved_phone[:3]}****{saved_phone[-4:]}\n")
                else:
                    phone_inputs = page.locator('input[placeholder*="手机"], input[placeholder*="电话"], input[type="tel"], input[name*="phone"], input[name*="mobile"]')
                    if await phone_inputs.count() > 0:
                        print("  📱 检测到手机号输入框，请输入手机号（将自动保存供下次使用）")
                        await phone_inputs.first.wait_for(state="visible", timeout=60000)
                        await asyncio.sleep(3)
                        phone_val = await phone_inputs.first.input_value()
                        if phone_val and len(phone_val) >= 11:
                            with open(PHONE_FILE, "w", encoding="utf-8") as f:
                                f.write(phone_val)
                            print(f"  ✅ 手机号已保存: {phone_val[:3]}****{phone_val[-4:]}\n")

            await page.wait_for_url('**/console/**', timeout=300000)
            print("✅ 登录成功，正在抓取套餐信息...")

            # 保存 Cookie
            cookies = await ctx.cookies()
            with open(COOKIE_FILE, 'w') as f:
                json.dump(cookies, f)

            # 抓取套餐信息
            plan_name = None
            plan_quota = None
            official_credits = None
            valid_until = None
            try:
                await asyncio.sleep(3)
                body_text = await page.inner_text('body')
                plan_match = re.search(r'(Lite|Standard|Pro|Max)\s*(月度|年度)\s*套餐', body_text, re.IGNORECASE)
                if plan_match:
                    plan_name = plan_match.group(1)
                    plan_type = "monthly" if "月度" in plan_match.group(2) else "yearly"
                    print(f"  📋 检测到套餐: {plan_name} {plan_match.group(2)}套餐")

                credits_match = re.search(
                    r'([\d,]+)\s*/\s*([\d,]+)\s*\n\s*已使用\s*([\d.]+)%', body_text)
                if credits_match:
                    official_credits = int(credits_match.group(1).replace(',', ''))
                    plan_quota = int(credits_match.group(2).replace(',', ''))
                    used_pct = credits_match.group(3)
                    print(f"  💰 当前套餐用量: {official_credits:,} / {plan_quota:,} Credits (已使用 {used_pct}%)")

                valid_match = re.search(r'有效期至\s*(\d{4}-\d{2}-\d{2})', body_text)
                if valid_match:
                    valid_until = valid_match.group(1)
                    print(f"  📅 有效期至: {valid_until}")

                if plan_name:
                    save_plan_info(plan_name, plan_quota, plan_type, official_credits, valid_until)
                    print(f"  ✅ 信息已保存")
                else:
                    print("  ⚠️  未能从页面识别套餐，将使用用量反推")
            except Exception as e:
                print(f"  ⚠️  抓取套餐信息失败: {e}")

            # 导出 xlsx
            print("\n📥 正在导出数据...")
            export_btns = page.locator('button:has-text("导出")')
            await export_btns.last.wait_for(state="visible", timeout=10000)
            async with page.expect_download(timeout=30000) as dl_info:
                await export_btns.last.click()
            download = await dl_info.value
            xlsx = os.path.join(dl, download.suggested_filename)
            await download.save_as(xlsx)
            print(f"✅ 已导出: {download.suggested_filename}\n")

            official_data = {
                "official_credits": official_credits,
                "plan_name": plan_name,
                "quota": plan_quota,
                "valid_until": valid_until,
            }
            analyze_xlsx(xlsx, official_data=official_data)

        except Exception as e:
            print(f"❌ {e}")
        finally:
            await browser.close()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 主入口
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async def default_run():
    """默认模式：自动获取最新官网数据 + 分析"""
    print("🔍 正在获取官网最新数据...\n")

    # 1. 抓取官网最新数据
    official_data = await fetch_official_data()
    if official_data:
        pc = official_data.get('official_credits', 0)
        print(f"  ✅ 官网数据: {pc:,} Credits" if pc else "  ⚠️  未获取到 Credits")
    else:
        print("  ⚠️  无法获取官网数据，将使用本地缓存")
        official_data = {}

    # 2. 找最新的 xlsx
    files = sorted(glob.glob(str(Path.home() / "Downloads" / "token_plan_usage_data_*.xlsx")),
                   key=os.path.getmtime, reverse=True)
    if not files:
        print("❌ 无导出文件，先运行: python mimo-usage-checker.py --login")
        return

    print(f"\n📁 分析: {Path(files[0]).name}\n")
    analyze_xlsx(files[0], official_data=official_data)


def main():
    args = sys.argv[1:]

    if "--login" in args:
        asyncio.run(login())
    elif "--xlsx" in args:
        i = args.index("--xlsx")
        if i + 1 < len(args):
            analyze_xlsx(args[i + 1])
        else:
            print("❌ 缺少文件路径")
    else:
        asyncio.run(default_run())


if __name__ == "__main__":
    main()
