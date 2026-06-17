# 富途证券年度账单 → 资本利得报告生成器

从富途证券年度账单 Excel 文件自动生成资本利得（Capital Gains）报告，支持股票、期权、基金的移动加权平均（MWA）成本计算。

---

## 3 步上手（不需要懂编程）

### 第 1 步：安装 Python

**macOS（苹果电脑）**

打开「终端」App（按 `Command + 空格`，输入 `终端`），粘贴运行：

```bash
brew install python
```

如果提示 `brew: command not found`，先去 https://brew.sh 安装 Homebrew，或访问 https://www.python.org/downloads/macos/ 下载安装包。

如果安装 openpyxl时提示错误 `python3 -m pip install openpyxl`，使用如下方法解决：

```
# 1. 创建虚拟环境（在当前目录下创建 venv 文件夹）
python3 -m venv venv

# 2. 激活虚拟环境
source venv/bin/activate

# 3. 安装 openpyxl
pip install openpyxl
```

**Windows**

1. 访问 https://www.python.org/downloads/
2. 点击 **Download Python 3.xx.x**
3. 下载后双击安装，**务必勾选**「Add Python to PATH」
4. 点击「Install Now」完成安装

### 第 2 步：准备账单文件

从富途 App 导出每年的年度账单（Excel 格式），放入脚本所在的同一个文件夹：

```
2023_年度账单_12345678.xlsx
2024_年度账单_12345678.xlsx
2025_年度账单_12345678.xlsx
```

### 第 3 步：双击运行

| 系统 | 操作 |
|------|------|
| macOS | 双击 `start.command` |
| Windows | 双击 `start.bat` |

脚本会自动：
1. 检查 Python 和依赖（没有就自动安装）
2. 读取所有年度账单
3. 生成 `capital_gains_report_v2.html`（浏览器打开查看）

如果出错，脚本会给出明确的错误提示和解决方法。

---

## 常见问题

**Q: 双击 start.command 提示"无法打开，因为无法验证开发者"？**

macOS 安全限制。解决方法：
1. 右键点击 `start.command`
2. 选择「打开」
3. 在弹出的对话框中点击「打开」

**Q: 报告生成后怎么看？**

双击 `capital_gains_report_v2.html`，用浏览器打开即可。其中：
- **已实现盈亏** — 当年卖出产生的实际盈亏
- **年度交易概况** — 年初/年末持仓、买卖金额汇总
- **缴税估算** — 按 20% 税率折算人民币税额
- **浮动盈亏** — 年末持仓按市价计算的未实现损益

**Q: 为什么有些股票显示"未处理资产进出"？**

富途账单中的 **DTC IN**（转入户）和 **Gift**（赠予）没有买入成本记录，脚本无法自动计算。需要在 `script_config.json` 的 `overrides` 中手工补充。详见下方「成本覆盖」章节。

**Q: 年度卖出金额和富途 App 显示的为什么不完全一样？**

脚本按 `成交金额` 统计卖出总额，与富途口径一致。微小差异来自汇率精度（脚本使用固定汇率，富途使用实际成交日汇率），通常在 1-2% 以内。

**Q: 账户升级/合并后历史交易怎么处理？**

脚本会自动合并所有检测到的账户。内部转账事件（Account Upgrade）会被自动跳过，不会重复计算。

---

## 配置文件（可选）

如果一切正常，不需要改任何配置。遇到以下情况才需要编辑 `script_config.json`：

### 场景 1：只想处理部分账户

如果你的富途有多个子账户，想只统计其中几个，填写 `target_accounts`：

```json
{
  "target_accounts": ["1001209838769859"]
}
```

留空或不填 → 自动检测所有账户。

### 场景 2：补充 DTC IN 成本基础

当从其他券商转入股票到富途时，富途账单不知道你的真实买入成本。需要在 `overrides` 中补充：

```json
{
  "overrides": [
    {
      "code": "BABA",
      "currency": "USD",
      "dtc_in": [
        {"date": "2022-07-12", "qty": "500", "price": "220", "note": "从XX券商转入"},
        {"date": "2023-08-01", "qty": "618", "price": "140", "note": "从YY券商转入"}
      ]
    }
  ]
}
```

每个 `dtc_in` 条目：
- `date` — 转入日期（富途账单中的到账日）
- `qty` — 转入数量
- `price` — 你实际的买入均价
- `note` — 备注（可选）

脚本会自动用这些数据重算种子持仓的平均成本，并把转入注入为买入交易。

### 场景 3：调整税率或汇率

```json
{
  "tax_rate": "0.20",
  "fx_rates": {
    "2023": "7.0827",
    "2024": "7.1884",
    "2025": "7.0288"
  }
}
```

`fx_rates` 是各年末 USD 对 CNY 的中间价，用于税务折算。港币按 USD/CNY ÷ 7.8 折算。

### 完整配置说明

| 字段 | 必填 | 说明 | 默认值 |
|------|------|------|--------|
| `target_accounts` | 否 | 账户号码列表，留空自动检测 | 自动检测 |
| `default_option_multiplier` | 否 | 期权乘数兜底 | `{"USD": "100", "HKD": "500"}` |
| `underlying_to_stock` | 否 | 港股期权代码映射 | 见下方 |
| `tax_rate` | 否 | 资本利得税率 | `0.20` |
| `fx_rates` | 否 | 年末 USD/CNY 汇率 | 见下方 |
| `hkd_peg_divisor` | 否 | 港币联系汇率除数 | `7.8` |
| `overrides` | 否 | DTC IN 成本覆盖 | `[]` |
| `output_html` | 否 | 输出 HTML 文件名 | `capital_gains_report_v2.html` |
| `output_json` | 否 | 输出 JSON 文件名 | `capital_gains_report_v2.json` |

**港股期权代码映射**（默认已配置）：

```json
{
  "underlying_to_stock": {
    "CMB": "3968",
    "MET": "3690",
    "TCH": "700"
  }
}
```

如果报告生成时提示某个港股期权代码找不到映射，在这里添加即可。

---

## 报告内容详解

### 已实现盈亏（按 币种/年/资产类型）

按币种（USD/HKD）、年份、资产类型（股票/期权/基金）分类汇总当年平仓实现的盈亏。

### 细分类别

按平仓原因细分：主动平仓、到期作废（EXP-NA）、被指派（ASS）、行权（EXR）、强平。

### 底层资产 Top 10

按绝对盈亏金额排序，显示贡献最大的 10 只股票/期权。

### 年度交易概况

| 字段 | 说明 |
|------|------|
| 年初持仓数/市值 | 上年度期末快照 |
| 买入笔数/金额 | 含申购、买入开仓、买入平仓 |
| 卖出笔数/金额 | 含赎回、卖出平仓、卖出开仓 |
| 已实现盈亏 | 当年平仓实现的 PnL |
| 年末持仓数/市值 | 当年期末快照 |
| 浮动盈亏 | 年末市价 vs 平均成本的未实现损益 |
| HKD 合计 | 所有币种按 USD×7.8 + HKD 直接汇总 |

### 年度缴税估算

按当年各币种正收益 × 税率 × 年末汇率，折算为人民币税额。亏损不抵税。

### 剩余未平仓持仓

年末仍持有的头寸（多头/空头），以及未匹配到的成本覆盖信息。

---

## 技术说明

### 文件结构

```
工作目录/
├── calc_capital_gains_v2.py      # 主脚本
├── script_config.json             # 配置文件
├── cost_basis_overrides.json      # (可选) 旧版配置，向后兼容
├── test_calc_capital_gains_v2.py  # 单元测试
├── start.command                  # macOS 双击启动
├── start.bat                      # Windows 双击启动
├── YYYY_年度账单_*.xlsx           # 富途导出的年度账单
├── capital_gains_report_v2.html   # 生成的 HTML 报告
└── capital_gains_report_v2.json   # 生成的 JSON 数据
```

### 计算逻辑

本脚本采用 **移动加权平均**（Moving Weighted Average, MWA）计算持仓成本：

- 每次买入时，按数量加权更新该股票的平均成本
- 每次卖出时，按平均成本计算盈亏
- 不同于 FIFO（先进先出），MWA 只维护一条持仓记录，平仓后剩余持仓的平均成本不变

### 单元测试

开发者可运行测试验证脚本正确性：

```bash
python3 -m unittest test_calc_capital_gains_v2 -v
```

---

## 许可证

本脚本为个人用途开发，按原样提供，使用者自行承担风险。
