# 富途证券年度账单 → 资本利得报告生成器

从富途证券年度账单 Excel 文件自动生成资本利得（Capital Gains）报告，支持股票、期权、基金的移动加权平均（MWA）成本计算。

## 功能

- **自动扫描**目录中的 `YYYY_年度账单_*.xlsx` 文件
- **移动加权平均**（Moving Weighted Average）计算持仓成本
- **支持多账户**自动合并（账户升级/转移场景）
- **期权特殊事件**处理：行权（EXR）、指派（ASS）、到期作废（EXP）
- **DTC IN / Gift 成本覆盖**：手动补充转入股票的真实成本基础
- **年度交易概况**：年初/年末持仓、买卖金额、已实现/浮动盈亏
- **税务估算**：按年末汇率折算 CNY 税额
- **输出 HTML 报表** + **JSON 结构化数据**

## 依赖

```bash
pip install openpyxl
```
pip 找不到时，确认已经安装了python3，然后执行：
```bash
python3 -m pip install openpyxl
```

Python 3.10+

## 文件结构

```
工作目录/
├── calc_capital_gains_v2.py    # 主脚本
├── script_config.json           # 配置文件（用户可调整）
├── cost_basis_overrides.json    # (可选) 旧版成本覆盖文件，向后兼容
├── test_calc_capital_gains_v2.py # 单元测试
├── YYYY_年度账单_*.xlsx         # 富途导出的年度账单
├── capital_gains_report_v2.html # 生成的 HTML 报告
└── capital_gains_report_v2.json # 生成的 JSON 数据
```

## 快速开始

### 1. 准备账单文件

从富途 App 导出每年的年度账单（Excel 格式），放入同一目录：

```
2023_年度账单_12345678.xlsx
2024_年度账单_12345678.xlsx
2025_年度账单_12345678.xlsx
```

### 2. 创建配置文件

复制 `script_config.json` 到工作目录，按需修改：

```json
{
  "target_accounts": [],
  "tax_rate": "0.20",
  "fx_rates": {
    "2023": "7.0827",
    "2024": "7.1884",
    "2025": "7.0288"
  },
  "overrides": []
}
```

**`target_accounts`**（可选）：
- 留空 → 脚本自动从 xlsx 中扫描所有账户号码
- 手动填写 → 只处理指定账户（用于过滤子账户）

### 3. 运行脚本

```bash
python3 calc_capital_gains_v2.py /path/to/bills
```

或在工作目录内直接运行：

```bash
python3 calc_capital_gains_v2.py .
```

### 4. 查看报告

- **HTML 报告**：用浏览器打开 `capital_gains_report_v2.html`
- **JSON 数据**：供其他程序消费 `capital_gains_report_v2.json`

## 配置文件详解（script_config.json）

| 字段 | 必填 | 说明 |
|------|------|------|
| `target_accounts` | 否 | 账户号码列表，留空自动检测 |
| `default_option_multiplier` | 否 | 期权乘数兜底，`{"USD": "100", "HKD": "500"}` |
| `underlying_to_stock` | 否 | 港股期权代码映射，如 `{"CMB": "3968"}` |
| `tax_rate` | 否 | 资本利得税率，默认 `0.20` |
| `fx_rates` | 否 | 年末 USD/CNY 汇率，用于税务折算 |
| `hkd_peg_divisor` | 否 | 港币联系汇率除数，默认 `7.8` |
| `overrides` | 否 | DTC IN / Gift 成本覆盖，见下方 |
| `skip_categories` | 否 | 要忽略的交易品类，默认 `[]` |
| `skip_directions` | 否 | 要忽略的交易方向，默认 `[]` |
| `output_html` | 否 | 输出 HTML 文件名 |
| `output_json` | 否 | 输出 JSON 文件名 |

## DTC IN / Gift 成本覆盖（Overrides）

当富途账单中的**转入户（DTC IN）**或**赠予（Gift）**没有真实买入价时，需要在 `overrides` 中手工补充：

```json
{
  "overrides": [
    {
      "code": "BABA",
      "currency": "USD",
      "dtc_in": [
        {"date": "2022-07-12", "qty": "500", "price": "220", "note": "DTC IN 2022-07-12"},
        {"date": "2023-08-01", "qty": "618", "price": "140", "note": "DTC IN 2023-08-01"}
      ]
    }
  ]
}
```

**处理逻辑**：
1. 用 `dtc_in` 列表（按时间排序）做 **MWA 重算**种子持仓的 `avg_cost`
2. 把 >= 首个持仓年份 的转入注入为**买入开仓 trade**
3. 如果同时存在旧版 `cost_basis_overrides.json`，其内容会**覆盖** `script_config.json` 中的对应字段

## 单元测试

```bash
python3 -m unittest test_calc_capital_gains_v2 -v
```

## 报告内容说明

### HTML 报告板块

1. **已实现盈亏（按 币种/年/资产类型）** — 股票/期权/基金分类汇总
2. **细分类别** — 主动平仓、到期作废、被指派、强平等
3. **底层资产 Top 10** — 按绝对值排序的盈亏贡献
4. **年度交易概况** — 年初/年末持仓、买卖金额、已实现/浮动盈亏、HKD 合计
5. **年度缴税估算** — 按 20% 税率、年末汇率折算 CNY
6. **期权特殊事件明细** — 行权/指派/到期的权利金处理
7. **剩余未平仓持仓** — 年末仍持有的头寸

### 年度交易概况字段

| 字段 | 说明 |
|------|------|
| 年初持仓数/市值 | 上年度期末快照 |
| 买入笔数/金额 | 含申购、买入开仓、买入平仓（平空头） |
| 卖出笔数/金额 | 含赎回、卖出平仓、卖出开仓 |
| 已实现盈亏 | 当年平仓实现的 PnL |
| 年末持仓数/市值 | 当年期末快照 |
| 浮动盈亏 | 年末市价 vs 平均成本的未实现损益 |
| HKD 合计 | 所有币种按 USD×7.8 + HKD 直接汇总 |

## 常见问题

**Q: 为什么年度卖出金额和富途 App 显示的不完全一致？**
A: 脚本按 `成交金额` 统计卖出总额，与富途口径一致。微小差异来自汇率精度（脚本使用固定 7.8 或配置汇率，富途使用实际成交日汇率）。

**Q: 为什么有些资产进出显示"未处理"？**
A: 脚本只处理 EXP（到期）、ASS（指派）、EXR（行权）三类期权事件。DTC IN / Gift 等需要通过 `overrides` 手工补充成本基础。

**Q: 账户升级/合并后历史交易怎么处理？**
A: 脚本会自动合并所有检测到（或配置）的账户，内部转账事件会被自然跳过。
