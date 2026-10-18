# 题目：贷款资产池回款与分配瀑布（Loan Pool Collection & Distribution Waterfall）

你是一家结构化信贷基金的量化分析师。该基金正在清盘一个贷款资产池，需要通过一个**多期、多档层（tranche）、且含期末 clawback（追回）**的分配瀑布，把回款分配给 LP（以及 GP 的 carried interest）。基于源工作簿 `/data/init.xlsx`，搭建完整模型并把答案写入 `/output/result.xlsx`。

下面所有规则都是**显式且完整的**——没有任何隐藏假设。请严格按所写实现。偏离既定顺序或计提基数都会导致勾稽不平，进而通不过一致性检查。

---

## 1. 输入

`/data/init.xlsx` 含两张表：

**`LoanTape`** —— 每笔贷款一行。列：
`LoanID, IssueDate, MaturityDate, FaceValue, AnnualRate, DefaultFlag, PaidPrincipal, PaidInterest, RecoveryRate, Fee`

**`Params`** —— 参数，以「键 / 值」行的形式给出：
- `FundStartDate` —— 例：`01-Jan-2024`。
- `DistributionDates` —— 用**分号分隔**的 K 个分配日列表（格式 `DD-Mon-YYYY`），例：`31-Dec-2025;31-Dec-2026;31-Dec-2027;31-Dec-2028`。
- `PreferredReturnRate` —— 年化优先回报利率（**单利**），例：`0.08`。
- `GPCarry` —— GP 合同约定的 carried interest，例：`0.15`（即 15%）。
- `ExitGainPct` —— 组合层面的已实现收益，表示为**总资本**的比例（该贷款组合被溢价出售）。例：`0.25` = +25%。
- `DayCount` —— `30/360`。
- `TrancheACapital`, `TrancheBCapital`, `TrancheCCapital` —— 各档资本，表示为**总资本的比例**（三者之和为 1.0）。A 为优先档（senior），B 为夹层档（mezzanine），C 为劣后档（junior）。

### 1.1 脏数据解析（重要）
- **Params 中的日期**（`FundStartDate`、`DistributionDates`）格式为 `DD-Mon-YYYY`。`DistributionDates` 以分号分隔；按 `;` 切分得到 K 个期。
- **`Fee`**（LoanTape）—— 由资产池承担的费用。它可能是**负数**（数据噪声）或**空白**。取绝对值；空白 → `0`。
- **`DefaultFlag`** 取值 `Y` / `N`。
- `IssueDate` / `MaturityDate` 出现在 tape 中（格式混杂），但**本模型不使用**——回款直接取自 `PaidPrincipal` / `PaidInterest`（见 §2）。仅在你想展示它们时才做防御性解析；它们从不进入计算。

---

## 2. 单笔贷款计算 → 表 `LoanSchedules`

对每笔贷款（每行一笔，**顺序与 tape 一致**）计算以下列：

`LoanID, Defaulted, CollectedPrincipal, CollectedInterest, Recovery, Fee, ChargeOff, NetToPool`

- `LoanID` —— 照原样。
- `Defaulted` —— 照原样的 `Y` / `N`。
- `Fee` —— `abs(解析后的 Fee)`，空白 → `0`（保留 2 位小数）。

正常贷款（`DefaultFlag = N`）：
- `CollectedPrincipal = PaidPrincipal`（已收回的现金本金，照记录）。
- `CollectedInterest = PaidInterest`（已收回的现金利息，照记录）。
- `Recovery = 0`，`ChargeOff = 0`。

违约贷款（`DefaultFlag = Y`）：
- `Recovery = round(FaceValue × RecoveryRate, 2)`。
- `CollectedPrincipal = round(PaidPrincipal + Recovery, 2)`。
- `CollectedInterest = PaidInterest`（违约前实际收到的利息）。
- `ChargeOff = round(FaceValue − Recovery, 2)`（仅备忘；不要在任何地方再次减它）。

两类贷款：
- `NetToPool = round(CollectedPrincipal + CollectedInterest − Fee, 2)`。

---

## 3. 资产池汇总 → 表 `PoolSummary`

双列表（`Metric, Value`），按以下顺序列出这些行：

`TotalCapital, TotalCollectedPrincipal, TotalCollectedInterest, TotalRecovery, TotalFees, TotalChargeOff, RealizedExitGain, DistributableCash`

- `TotalCapital` = Σ `FaceValue`（全部贷款）。
- `TotalCollectedPrincipal` = Σ `CollectedPrincipal`。
- `TotalCollectedInterest` = Σ `CollectedInterest`。
- `TotalRecovery` = Σ `Recovery`。
- `TotalFees` = Σ `Fee`。
- `TotalChargeOff` = Σ `ChargeOff`。
- `RealizedExitGain = ExitGainPct × TotalCapital`（组合层面）。
- `DistributableCash = TotalCollectedPrincipal + TotalCollectedInterest − TotalFees + RealizedExitGain`。
  （`TotalRecovery` 对违约贷款已计入 `CollectedPrincipal` 内，`TotalChargeOff` 仅为备忘——二者都不在此再次加减。）

---

## 4. 分配瀑布 → 表 `Waterfall`（+ `WaterfallDetail`）

### 4.0 资本结构与准备
共有**三个 LP 档**，其资本合计等于 `TotalCapital`：
- 档 **A**（优先）：cap `A = TrancheACapital × TotalCapital`
- 档 **B**（夹层）：cap `B = TrancheBCapital × TotalCapital`
- 档 **C**（劣后）：cap `C = TrancheCCapital × TotalCapital`

GP 不出资，仅赚取 carried interest。

令 `D = DistributableCash`，`K =` 分配日个数。每期 `k = 1..K` 获得相等的份额 `slice_k = D / K`，并运行下面**完整**的瀑布。每个档的未返还资本（以及 GP 的累计 carry）**跨期结转**（有状态级联）。

跨期结转的量：
- `ret_t` = 已返还给档 `t` 的资本（期初为 0）。
- `prefpaid_t` = 已支付给档 `t` 的优先回报（期初为 0）。
- `gp_running` = 迄今累计支付的 GP carry（期初为 0）。
- `prev_1 = FundStartDate`；当 `k>1` 时，`prev_k =` 上一期分配日。
- 年分数 `yf_k = days360(prev_k, dist_k) / 360`（30/360 计日法）。

### 4.1 每期瀑布（对 k = 1..K 按日期顺序各运行一次）

**第 1 步 —— 优先回报计提（本期应计额）。** 对每档 `t`，按其**期初未返还资本**计提：
`accrued_t = (cap_t − ret_t) × PreferredReturnRate × yf_k`
`PrefDue_k = Σ_t accrued_t`。
（单利，不复合。未支付的应计优先回报**不**结转至后期——它不会被加入任何未来的计提基数。）

**第 2 步 —— 返还资本（RoC）**给 LP 各档，按优先顺序 **A → B → C**：
对 `t` 依次取 `[A, B, C]`：
`give_t = min(cap_t − ret_t, avail)`；`ret_t += give_t`；`avail −= give_t`。

**第 3 步 —— 支付优先回报。** `pay_pref = min(PrefDue_k, avail)`。
按 `accrued_t` 的**比例**分摊 `pay_pref`：
`prefpaid_t += round(pay_pref × accrued_t / PrefDue_k, 2)`（若 `PrefDue_k > 0`；否则为 `0`）。`avail −= pay_pref`。

**第 4 步 —— GP catch-up。** 令 `cum_lp_gross = Σ_t ret_t + Σ_t prefpaid_t`（迄今累计的 LP RoC + 优先回报）。GP 的目标累计 carry 为 `GPCarry / (1 − GPCarry) × cum_lp_gross`。支付缺口：
`gp_due = max(0, target − gp_running)`；`give_catchup = min(gp_due, avail)`；`gp_running += give_catchup`；`avail −= give_catchup`。

**第 5 步 —— 剩余分配。** `gp_excess_k = round(GPCarry × avail, 2)`；`lp_excess_k = round(avail − gp_excess_k, 2)`；`avail = 0`。

### 4.2 汇总与期末 clawback
K 期全部跑完后：
- `ReturnOfCapital = Σ_t ret_t`（返还给 LP 的总资本）。
- `PreferredReturn = Σ_t prefpaid_t`。
- `GPCatchUp = gp_running`（累计支付的 catch-up）。
- `LPExcess = Σ_k lp_excess_k`，`GPExcess = Σ_k gp_excess_k`。

**Clawback（期末清算）。** 计算 LP 在整个基金生命周期内的完整合同权益：
`full_pref = TotalCapital × PreferredReturnRate × yf_total`，
其中 `yf_total = days360(FundStartDate, last_distribution_date) / 360`。
`lp_entitlement = TotalCapital + full_pref`。
若**同时满足** `ReturnOfCapital + PreferredReturn < lp_entitlement` **且** `D < lp_entitlement`：
`Clawback = min(GPCatchUp + GPExcess, lp_entitlement − (ReturnOfCapital + PreferredReturn))`
否则 `Clawback = 0`。
（被追回的金额由 GP 返还给 LP。）

汇总：
- `GPTotal = GPCatchUp + GPExcess − Clawback`。
- `LPTotal = D − GPTotal`（精确到分；等价地，`LPTotal = ReturnOfCapital + PreferredReturn + LPExcess + Clawback`）。

### 4.3 表 `Waterfall`（双列：`Metric, Value`），按此顺序列出各行：
`Distributions` (= K),
`TrancheACapital`, `RoC_TrancheA`, `Pref_TrancheA`,
`TrancheBCapital`, `RoC_TrancheB`, `Pref_TrancheB`,
`TrancheCCapital`, `RoC_TrancheC`, `Pref_TrancheC`,
`ReturnOfCapital`, `PreferredReturn`, `GPCatchUp`, `LPExcess`, `GPExcess`,
`Clawback`, `LPTotal`, `GPTotal`, `SumCheck`。

（`RoC_TrancheX` = 各期累计返还给档 X 的资本；`Pref_TrancheX` = 各期累计支付给档 X 的优先回报；`SumCheck = LPTotal + GPTotal`。）

### 4.4 表 `WaterfallDetail`（每期追溯），列：
`Distribution, RoC_A, RoC_B, RoC_C, Pref_A, Pref_B, Pref_C, GPCatchUp, LPExcess, GPExcess`
每期一行（`D1, D2, … Dk`），填入该期的数值。它展示有状态级联，满分必需。

---

## 5. 输出规则
- 写出 `LoanSchedules`、`PoolSummary`、`Waterfall`、`WaterfallDetail` 四张表，使用上面列出的**确切名称**与**确切的 metric / 列标签**。
- **表名与 Metric / 列标签必须使用上述英文原文**（判卷按原文精确匹配，写成中文会判错）；其余说明文字可中文。
- 报告时每个金额四舍五入到**分（2 位小数）**。
- 工作簿必须**精确到分地内部一致**：
  - `DistributableCash = TotalCollectedPrincipal + TotalCollectedInterest − TotalFees + RealizedExitGain`；
  - `LPTotal + GPTotal = DistributableCash`；
  - `ReturnOfCapital ≤ TotalCapital`；
  - 对每档 X，`RoC_TrancheX ≤ TrancheXCapital`；
  - `GPTotal = GPCatchUp + GPExcess − Clawback`；
  - §4.2 的 clawback 逻辑成立（当 LP 不足本金 + 全期优先回报、且资产池本身也不足时，GP 吐回 carry）。

用 Python + `openpyxl`（已安装）。读取 `/data/init.xlsx`，写入 `/output/result.xlsx`。
