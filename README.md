# ETF轮动策略

股票型ETF的动态指数轮动研究项目。主要流程为：

1. 月度ETF池筛选；
2. 指数收益率聚类去重；
3. 计算两种趋势因子；
4. 选择排名前1%的指数，并在排名后过滤窗口收益率不大于0的指数；
5. 映射至成交量最大的跟踪ETF；
6. 按收盘价或次日VWAP回测，并计入买卖各0.1%的交易成本。

## 当前基线参数

- 聚类相关性阈值：0.8
- 趋势窗口：40个交易日
- 入选比例：前1%（`TOP_PERCENT = 0.010`）
- 调仓模式：11个账户逐日错峰，每个账户持有11个交易日
- 输出目录：`staggered_11_accounts_hold_11d`

回测脚本同时支持：

- `REBALANCE_MODE = "rebalance"`：输出到 `rebalance_<x>d`；
- `REBALANCE_MODE = "staggered"`：输出到 `staggered_<x>_accounts_hold_<x>d`。

其中 `x` 由 `ACCOUNT_REBALANCE_INTERVAL` 设置；账户数量会自动计算。

## 项目结构

```text
scripts/
  ETF数据下载/
  ETF池筛选/
  ETF趋势策略回测/
  ETF策略检验/
outputs/
  etf_data/                # ETF原始数据和派生宽表
  etf_pool/                # ETF池与指数聚类结果
  etf_strategy_test/       # RankIC等检验结果
  etf_trend_strategy/      # 因子、指数行情和全部回测输出
```

## 数据与凭据

GitHub版本完整保存 `scripts/`、`outputs/` 和本说明文件；`outputs/` 中的大型文件通过Git LFS保存。

iFinD凭据只保存在项目根目录唯一的 `.env` 文件中，下载脚本启动时会自动读取：

```text
IFIND_USERNAME=...
IFIND_PASSWORD=...
IFIND_REFRESH_TOKEN=...
```

`.env` 已被Git忽略，不会提交到GitHub。
