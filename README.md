# subitem-status-tracker / 子项状态统计器

> 零依赖的 VBA 宏：把一堆「子项」（箱号、批次、流水号、序列号……）按主单号归组，自动统计「已处理 / 未处理」，一键生成带高亮、边框、筛选和冻结表头的汇总表。
> A dependency-free VBA macro that groups sub-items under a master number, tallies processed vs. unprocessed, and builds a one-click summary sheet.

> *非官方项目。与 WPS Office（金山）、Microsoft（VBA）无隶属、赞助或背书关系；提及产品名仅用于说明运行环境。*

---

## 痛点：它解决什么问题（中文）

几乎所有「主单号—子项」场景，最后都卡在同一个动作上：**拿到一份子项清单，快速知道「这个主单号下面总共多少、已完成多少、还剩多少没处理」**。

典型场景：

- **仓库拣货 / 入库**：一张采购单（主单号）拆成几十个箱（子项），要立刻看出哪些箱还没拣、哪些已拣。
- **物流清单对账**：一批运单号下面挂着若干件/序列号，跟单员需要一眼看出哪些运单还差几件没回传。
- **生产批次追踪**：一个批次号对应若干生产子批，质量或进度统计要按批次聚合。
- **订单履约 / 库存质检**：订单号下面的明细行，哪些已处理、哪些还挂着。

**手工做法的崩溃点**：在 Excel 里筛选→计数→条件格式标红→复制成汇总表。几十个主单号还能忍，几百个就必然出错、漏看、格式乱、下次还得重做。更糟的是，这类表每天/每周都在变，重复劳动没有尽头。

**这个宏的价值**：把「归组 + 计数 + 异常高亮 + 一键汇总」固化成一个宏。你只要保证两列数据——

- 一列是**子项编码**（如 `PO2025001-001`，主单号在分隔符之前）；
- 一列是**关联引用**（子项处理完就填、没处理就留空）；

宏会自动从子项编码里解析出主单号、按主单号聚合总数/已处理/未处理，把**未处理数量用浅红高亮**，并生成一张带粗体表头、自适应列宽、边框、自动筛选和冻结首行的干净汇总表。

**谁适合用**：仓管、物流跟单、生产统计、订单处理、库存 QC——任何每天和「主单号—子项」表格打交道、又不想写代码的人。改几行顶部常量就能套到你自己的表结构。

## What problem it solves (English)

Manually tallying sub-items under master numbers (boxes under a PO, serials under a shipment, batches under a lot) is slow and error-prone once you have hundreds of groups. This macro automates grouping, counting processed/unprocessed, red-highlighting exceptions, and produces a clean, filterable, frozen-header summary sheet — all in pure VBA with zero add-ins. Point it at two columns and run one macro.

---

## 快速运行（中文）

**环境要求**

- WPS Office（需支持 VBA）**或** Windows 版 Microsoft Excel；
- 一张数据表：一列放**子项编码**，另一列放**关联引用**（处理完有值、未处理留空）。

**三步跑起来**

1. 用 WPS / Excel 打开你的工作簿，按 `Alt + F11` 打开 VBA 编辑器；
2. `文件 ▸ 导入文件…`，选择 `src/SubItemTally.bas`；
3. 关闭编辑器，按 `Alt + F8` 运行宏 **`TallySubItems`**。会生成（或刷新）一张 `Summary` 汇总表。

**想立刻看到效果**：打开 `examples/sample-data.xlsx`，直接运行即可。默认列：子项编码在 **A 列**、关联引用在 **D 列**、表头在第 1 行、数据从第 2 行开始。

**套到自己的表**：改 `SubItemTally.bas` 顶部的常量（分隔符、列号、表名、是否只列已完成的）即可，详见 `docs/Configuration.md`。

## Quick start (English)

1. Open your workbook in WPS/Excel, press `Alt+F11`, `File ▸ Import File…` → `src/SubItemTally.bas`.
2. Press `Alt+F8` and run **`TallySubItems`**.
3. Open `examples/sample-data.xlsx` to try it instantly (default: sub-item code in A, linked reference in D). Tune the constants at the top of the module for your own layout.

---

## 工作原理 / How it works

1. 读取源表每一行（子项编码一列、关联引用一列）；
2. 从子项编码中按分隔符截取**主单号**（分隔符之前的部分）；
3. 用两个字典分别记录「主单号集合」和「每个子项的处理状态」；
4. 对每个主单号汇总 总数 / 已处理 / 未处理，写一行到汇总表；
5. 格式化（高亮、边框、筛选、冻结）并报完成。

> **归组规则**：主单号按子项编码的**前缀**匹配。主单号之间不能互为前缀（例如 `PO1` 和 `PO10` 会冲突）。请用互不重叠的编号规则，或调整分隔符让切分无歧义。

## 配置要点 / Configuration

所有开关都是模块顶部的常量（列号、分隔符、表名、输出规则），详见 [`docs/Configuration.md`](docs/Configuration.md)。**参数名唯一真源是 `src/SubItemTally.bas` 中的常量**，修改代码后请同步此处。

## 适用场景 / Use cases

仓库拣货、物流清单对账、生产批次追踪、订单履约核对、库存质检，以及任何「主单号挂子项、要统计完成度」的表格工作。

## 法律声明 / Legal & Non-affiliation

本仓库是 **ivy92888-glitch** 独立、clean-room 重写的通用工具，不含任何前雇主或第三方的专有数据、截图或业务特定内容。MIT 许可证仅覆盖作者本人的实现，不代表或转让任何第三方的权利。产品名（WPS Office、Microsoft Excel）为其各自所有者商标，仅用于说明运行环境。使用者需确保自身数据处理符合所在地法律与所在单位的政策。

## License

[MIT License](LICENSE) · Copyright (c) 2026 ivy92888-glitch.

---

<details><summary>维护说明（中英双语单文件，怎么避免漂移）</summary>

本文档是**单文件中英双语**：中文为权威详细描述，英文为简洁摘要（不强求逐句对等）。所有参数名（如 `KEY_COL`）直接取自源码常量，唯一真源是 `src/*.bas`。修改代码逻辑时，只要同步更新源码常量名与本文档「配置要点」中的同名常量即可，叙述性文字按需增删，不影响一致性。

</details>
