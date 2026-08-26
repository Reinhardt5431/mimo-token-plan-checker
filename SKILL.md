---
name: mimo-token-plan-checker
description: "MiMo Token Plan 用量分析：导出xlsx并用官方规则换算Token到Credits。"
version: 1.0.0
author: hermes-agent
---

# MiMo Token Plan 用量分析工具

分析小米 MiMo Token Plan 的消耗情况，用官方换算规则将 Token 折算为 Credits，对比套餐额度。

## 触发条件

当用户提到以下关键词时激活：
- MiMo 用量、Token Plan、Credits 消耗、套餐用量
- mimo-usage-checker、分析用量、查看消耗

## 使用方法

### 首次使用：登录并导出

```bash
pip install playwright openpyxl && playwright install chromium
python <skill_dir>/scripts/mimo-usage-checker.py --login
```

脚本会弹出浏览器，用户正常登录（支持扫码/短信/密码），登录成功后**自动从页面抓取套餐名称和额度**，保存到 `~/.mimo-usage-checker/plan_info.json`，然后导出 xlsx 并分析。

### 后续使用：分析已有文件

```bash
# 自动查找最近的导出文件（使用已保存的套餐信息）
python <skill_dir>/scripts/mimo-usage-checker.py

# 指定文件
python <skill_dir>/scripts/mimo-usage-checker.py --xlsx <file.xlsx>

# 手动指定套餐（覆盖 plan_info.json）
python <skill_dir>/scripts/mimo-usage-checker.py --plan standard
```

### 套餐识别优先级

1. `--plan` 命令行参数（最高优先级）
2. `~/.mimo-usage-checker/plan_info.json`（login 时自动保存）
3. 用量反推（兜底，会标注"推测"，不推荐）

## 官方换算规则（Credits per Token）

| 模型 | 输入（命中缓存） | 输入（未命中缓存） | 输出 |
|------|----------------|-------------------|------|
| mimo-v2.5 | 2 | 100 | 200 |
| mimo-v2.5-pro | 2.5 | 300 | 600 |
| mimo-v2.5-asr | 30M Credits/小时音频 | - | - |
| mimo-v2.5-tts 系列 | 限时免费 | - | - |

## 套餐额度（官方固定）

| 套餐 | 月度 | 年度 |
|------|------|------|
| Lite | 4.1B | 49.2B |
| Standard | 11B | 132B |
| Pro | 38B | 456B |
| Max | 82B | 984B |

## 报告内容

1. **汇总**：Token 总消耗、Credits 总消耗、套餐使用率
2. **按模型统计**：每个模型的 Token/Credits/请求占比
3. **每日明细**：每天的 Token 明细（输入命中/未命中/输出）+ 对应 Credits
4. **缓存命中率分析**：未命中缓存是 Credits 消耗的主要驱动因素
5. **非高峰优惠提示**：北京时间 00:00-08:00 享受 0.8x 系数

## 注意事项

- **xlsx 导出包含全部历史数据，管理台只统计当前订阅周期内的 Credits**。两者 Token 总数一致说明换算公式正确，Credits 差异纯粹是时间范围不同。分析时应以管理台显示的 Credits 为准，xlsx 的 Credits 仅作参考。
- 非高峰期（北京时间 00:00-08:00）消耗系数为 0.8x
- TTS 系列模型限时免费，不消耗 Credits
- 导出的 xlsx 文件在 `~/Downloads/token_plan_usage_data_*.xlsx`
- Cookie 保存在 `~/.mimo-usage-checker/cookies.json`
- 套餐信息保存在 `~/.mimo-usage-checker/plan_info.json`（login 时自动抓取）
