---
name: mimo-token-plan-checker
description: "MiMo Token Plan 用量分析：每次自动获取官网最新Credits，与xlsx换算结果交叉验证。"
version: 2.0.0
author: hermes-agent
---

# MiMo Token Plan 用量分析工具

分析小米 MiMo Token Plan 的消耗情况。**每次运行自动获取官网最新 Credits**，与 xlsx 导出数据换算结果交叉验证。

## 触发条件

当用户提到以下关键词时激活：
- MiMo 用量、Token Plan、Credits 消耗、套餐用量
- mimo-usage-checker、分析用量、查看消耗

## 使用方法

### 日常使用（推荐）

```bash
python <skill_dir>/scripts/mimo-usage-checker.py
```

自动执行：
1. 用已有 Cookie **无头浏览器**获取官网最新 Credits
2. 找到最近的 xlsx 导出文件
3. 按官方换算规则计算 Credits
4. **强制对比**：xlsx 计算值 vs 官网实时值
5. 输出完整分析报告

### 重新登录（Cookie 过期时）

```bash
python <skill_dir>/scripts/mimo-usage-checker.py --login
```

弹出浏览器，自动切换到手机号验证，登录后自动导出并分析。

### 分析指定文件

```bash
python <skill_dir>/scripts/mimo-usage-checker.py --xlsx <file.xlsx>
```

## 分析报告内容

1. **① 换算一致性验证**（核心）：xlsx 计算值 vs 官网实时 Credits，自动按订阅周期过滤历史数据
2. **② 每日消耗分析**：订阅周期内的每日 Token/Credits 明细 + 趋势分析
3. **③ 按模型汇总**：各模型的 Token/Credits/请求占比
4. **④ 套餐用量预估**：基于官网实时数据，按实际订阅周期预估用量

## 官方换算规则

| 模型 | 输入（命中缓存） | 输入（未命中缓存） | 输出 |
|------|----------------|-------------------|------|
| mimo-v2.5 | 2 Credits | 100 Credits | 200 Credits |
| mimo-v2.5-pro | 2.5 Credits | 300 Credits | 600 Credits |
| mimo-v2.5-asr | 30M Credits/小时 | - | - |

## 套餐额度

| 套餐 | 月度 | 年度 |
|------|------|------|
| Lite | 4.1B | 49.2B |
| Standard | 11B | 132B |
| Pro | 38B | 456B |
| Max | 82B | 984B |

## 注意事项

- **每次运行都会自动获取最新官网数据**，确保对比有效
- Cookie 保存在 `~/.mimo-usage-checker/cookies.json`，过期后需 `--login` 重新登录
- 手机号保存在 `~/.mimo-usage-checker/phone.txt`，首次登录后自动保存
- 套餐信息保存在 `~/.mimo-usage-checker/plan_info.json`
- 导出的 xlsx 在 `~/Downloads/token_plan_usage_data_*.xlsx`

## License

MIT
