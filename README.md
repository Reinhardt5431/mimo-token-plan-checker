# MiMo Token Plan 用量分析工具

分析小米 MiMo Token Plan 的消耗情况，用官方换算规则将 Token 折算为 Credits，对比套餐额度。

## 功能

- 📊 **总览**：Token/Credits 消耗、套餐使用率
- 📅 **每日明细**：每天的 Token 明细 + 缓存命中率
- 🤖 **按模型统计**：每个模型的 Token/Credits/请求占比
- 🎯 **套餐自动识别**：login 时自动抓取套餐名称和额度
- 💡 **非高峰优惠提示**：北京时间 00:00-08:00 享受 0.8x 系数

## 安装

```bash
pip install playwright openpyxl
playwright install chromium
```

## 使用

### 首次使用：登录并导出

```bash
python mimo-usage-checker.py --login
```

脚本会弹出浏览器，正常登录 MiMo 控制台（支持扫码/短信/密码），登录成功后自动抓取套餐信息并导出分析。

### 分析已有文件

```bash
# 自动查找最近的导出文件
python mimo-usage-checker.py

# 指定文件
python mimo-usage-checker.py --xlsx <file.xlsx>

# 手动指定套餐
python mimo-usage-checker.py --plan standard
```

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

## License

MIT
