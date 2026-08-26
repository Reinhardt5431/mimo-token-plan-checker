"""
MiMo Token Plan 用量分析工具

用法：
  python mimo-usage-checker.py --login              # 首次登录（弹出浏览器）
  python mimo-usage-checker.py                      # 分析最近的导出文件
  python mimo-usage-checker.py --xlsx <file.xlsx>   # 分析指定文件
  python mimo-usage-checker.py --plan standard      # 手动指定套餐
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


def save_plan_info(plan_name, quota, plan_type="monthly"):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    info = {"plan_name": plan_name, "quota": quota, "plan_type": plan_type}
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


def analyze_xlsx(filepath, plan_override=None, quota_override=None):
    try:
        import openpyxl
    except ImportError:
        print("❌ 需要安装: pip install openpyxl")
        return

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

    for row in ws.iter_rows(min_row=2, values_only=True):
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

    # 全局汇总
    total_tk = sum(v["total_tokens"] for v in daily.values())
    total_ih = sum(v["input_hit"] for v in daily.values())
    total_im = sum(v["input_miss"] for v in daily.values())
    total_ot = sum(v["output"] for v in daily.values())
    total_cr = sum(v["credits"] for v in daily.values())
    total_rc = sum(v["requests"] for v in daily.values())
    total_input = total_ih + total_im

    # 模型汇总
    model_agg = defaultdict(lambda: {
        "tokens": 0, "input_hit": 0, "input_miss": 0,
        "output": 0, "credits": 0, "requests": 0
    })
    for v in daily.values():
        for model, md in v["models"].items():
            for k in model_agg[model]:
                model_agg[model][k] += md[k]

    # ===== 输出 =====
    print()
    print("=" * 90)
    print("  MiMo Token Plan 用量分析")
    print("=" * 90)
    print()

    # --- 1. 总览 ---
    print("━━━ 1. 总览 ━━━")
    print()
    print(f"  套餐: {plan_name.upper()} | 额度: {fmt(quota)} Credits")
    print()
    print(f"  Token 总消耗:  {total_tk:>15,}")
    print(f"  缓存命中:      {total_ih:>15,}  ({total_ih/total_input*100:.1f}%)" if total_input else "")
    print(f"  缓存未命中:    {total_im:>15,}  ({total_im/total_input*100:.1f}%)" if total_input else "")
    print(f"  输出 token:    {total_ot:>15,}")
    print(f"  总请求数:      {total_rc:>15,}")
    print(f"  Credits 消耗:  {total_cr:>15,.0f}  (xlsx 计算)")
    print(f"  套餐额度:      {quota:>15,}")
    print(f"  使用率:        {total_cr/quota*100:>14.1f}%")
    print(f"  剩余额度:      {quota - total_cr:>15,.0f}")
    print()
    print("  ⚠️  注意: xlsx 导出包含全部历史数据，管理台只统计当前订阅周期内的 Credits。")
    print("     两者 Token 总数一致说明换算公式正确，Credits 差异纯粹是时间范围不同。")
    print("     以管理台显示的 Credits 为准。")
    print()

    # --- 2. 每日明细 ---
    print("━━━ 2. 每日明细 ━━━")
    print()
    print(f"  {'日期':<12} {'周':>2} {'Tokens':>12} {'缓存命中':>8} {'未命中':>10} {'输出':>8} {'Credits':>14} {'请求':>6}")
    print(f"  {'─'*12} {'─'*2} {'─'*12} {'─'*8} {'─'*10} {'─'*8} {'─'*14} {'─'*6}")

    for d in sorted(daily.keys()):
        v = daily[d]
        dt = datetime.datetime.strptime(d, "%Y-%m-%d")
        wd = WEEKDAYS[dt.weekday()]
        ti = v["input_hit"] + v["input_miss"]
        hit_pct = f"{v['input_hit']/ti*100:.0f}%" if ti else "—"
        print(f"  {d:<12} {wd:>2} {v['total_tokens']:>12,} {hit_pct:>8} {v['input_miss']:>10,} {v['output']:>8,} {v['credits']:>14,.0f} {v['requests']:>6}")

    print(f"  {'─'*12} {'─'*2} {'─'*12} {'─'*8} {'─'*10} {'─'*8} {'─'*14} {'─'*6}")
    hit_pct_all = f"{total_ih/total_input*100:.0f}%" if total_input else "—"
    print(f"  {'合计':<12} {'':>2} {total_tk:>12,} {hit_pct_all:>8} {total_im:>10,} {total_ot:>8,} {total_cr:>14,.0f} {total_rc:>6}")
    days = len(daily)
    print(f"  {'日均':<12} {'':>2} {total_tk//days:>12,} {'':>8} {'':>10} {'':>8} {total_cr/days:>14,.0f} {total_rc//days:>6}")
    print()

    # --- 3. 按模型汇总 ---
    print("━━━ 3. 按模型汇总 ━━━")
    print()
    print(f"  {'模型':<18} {'Tokens':>14} {'命中':>12} {'未命中':>10} {'输出':>10} {'Credits':>14} {'占比':>6} {'请求':>6}")
    print(f"  {'─'*18} {'─'*14} {'─'*12} {'─'*10} {'─'*10} {'─'*14} {'─'*6} {'─'*6}")

    for model, s in sorted(model_agg.items(), key=lambda x: -x[1]["credits"]):
        p = s["credits"] / total_cr * 100 if total_cr else 0
        print(f"  {model:<18} {s['tokens']:>14,} {s['input_hit']:>12,} {s['input_miss']:>10,} {s['output']:>10,} {s['credits']:>14,.0f} {p:>5.1f}% {s['requests']:>6}")

    print()
    print("  换算规则:")
    print("    mimo-v2.5:     命中×2    未命中×100   输出×200")
    print("    mimo-v2.5-pro: 命中×2.5  未命中×300   输出×600")
    print("    非高峰(00:00-08:00 北京) 0.8x 系数")
    print("=" * 90)
    print()


async def login():
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("❌ pip install playwright && playwright install chromium")
        sys.exit(1)

    os.makedirs(CONFIG_DIR, exist_ok=True)
    dl = str(Path.home() / "Downloads")

    print("🚀 浏览器即将弹出，请登录（支持扫码/短信/密码）...\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        ctx = await browser.new_context(accept_downloads=True)
        page = await ctx.new_page()
        await page.goto('https://platform.xiaomimimo.com/console/plan-manage')

        try:
            await page.wait_for_url('**/console/**', timeout=300000)
            print("✅ 登录成功，正在抓取套餐信息...")

            cookies = await ctx.cookies()
            with open(os.path.join(CONFIG_DIR, "cookies.json"), 'w') as f:
                json.dump(cookies, f)

            # 从页面抓取套餐信息
            plan_name = None
            plan_quota = None
            try:
                body_text = await page.inner_text('body')
                plan_match = re.search(r'(Lite|Standard|Pro|Max)\s*(月度|年度)\s*套餐', body_text, re.IGNORECASE)
                if plan_match:
                    plan_name = plan_match.group(1)
                    plan_type = "monthly" if "月度" in plan_match.group(2) else "yearly"
                    print(f"  📋 检测到套餐: {plan_name} {plan_match.group(2)}套餐")

                quota_match = re.search(r'[\d,]+\s*/\s*([\d,]+)', body_text)
                if quota_match:
                    plan_quota = int(quota_match.group(1).replace(',', ''))
                    print(f"  📊 套餐额度: {plan_quota:,} Credits")

                if plan_name:
                    save_plan_info(plan_name, plan_quota, plan_type)
                    print(f"  ✅ 套餐信息已保存到 {PLAN_INFO_FILE}")
                else:
                    print("  ⚠️  未能从页面识别套餐，将使用用量反推")
            except Exception as e:
                print(f"  ⚠️  抓取套餐信息失败: {e}")

            # 导出 xlsx
            print("\n📥 正在导出数据...")
            async with page.expect_download(timeout=30000) as dl_info:
                await page.click('button:has-text("导出")')
            download = await dl_info.value
            xlsx = os.path.join(dl, download.suggested_filename)
            await download.save_as(xlsx)
            print(f"✅ 已导出: {download.suggested_filename}\n")

            if plan_name and plan_quota:
                analyze_xlsx(xlsx, plan_override=plan_name, quota_override=plan_quota)
            else:
                analyze_xlsx(xlsx)

        except Exception as e:
            print(f"❌ {e}")
        finally:
            await browser.close()


def main():
    args = sys.argv[1:]

    if "--login" in args:
        asyncio.run(login())
    elif "--xlsx" in args:
        i = args.index("--xlsx")
        analyze_xlsx(args[i + 1]) if i + 1 < len(args) else print("❌ 缺少文件路径")
    else:
        files = sorted(glob.glob(str(Path.home() / "Downloads" / "token_plan_usage_data_*.xlsx")),
                       key=os.path.getmtime, reverse=True)
        if files:
            print(f"📁 {Path(files[0]).name}")
            analyze_xlsx(files[0])
        else:
            print("❌ 无导出文件，先运行: python mimo-usage-checker.py --login")


if __name__ == "__main__":
    main()
