# SAP MD04 BANF Real Demand Calculator
# SAP MD04 BANF 真实需求计算工具

A lightweight Python tool that analyzes SAP MD04 Excel exports and calculates how much of a selected purchase requisition (BANF) is actually required before a user-defined cutoff date.

一个用于分析 SAP MD04 导出数据的 Python 工具，可根据用户指定的截止日期计算某张采购申请（BANF）的真实需求量。

---

## Why This Tool? | 为什么开发这个工具？

In SAP MD04, long-term purchase requisitions are often generated automatically by MRP.

For buyers, the most important question is usually not:

> How much demand exists for this material?

Instead, the key question is:

> How much of this specific BANF is really needed before a specific date?

在 SAP MD04 中，MRP 会自动生成大量采购申请。

对于采购人员来说，最重要的问题通常不是：

> 这个物料一共有多少需求？

而是：

> 截止某个日期前，这张 BANF 实际需要多少数量？

This tool is designed to answer exactly that question.

本工具专门用于回答这个问题。

---

## Business Logic | 业务逻辑

This project uses a BANF-oriented calculation method.

本项目采用“面向 BANF”的计算方式。

Instead of recalculating the complete material balance from the top of MD04, the tool uses the available quantity immediately before the selected BANF and evaluates the movements after that BANF until the cutoff date.

本工具不是从 MD04 最顶部重新计算整个物料库存，而是以目标 BANF 前一行的可用数量为起点，评估目标 BANF 之后至截止日期之间的需求和其他收货。

### Formula | 计算公式

```text
Real BANF Requirement
=
Demand after BANF until cutoff date
- Available quantity before BANF
- Other receipts after BANF until cutoff date
```

对应中文：

```text
BANF 真实需求量
=
BANF 之后至截止日期的需求
- BANF 之前的可用数量
- BANF 之后至截止日期的其他收货
```

The result is limited to a minimum of zero and a maximum of the original BANF quantity.

计算结果最低为 0，最高不超过目标 BANF 的原始数量。

```text
Required BANF Quantity
= min(
    Original BANF Quantity,
    max(0, Calculated Requirement)
  )
```

### Safety Stock | 安全库存

The available quantity shown immediately before the target BANF is taken directly from the MD04 column `Verfügb. Menge`.

程序直接读取目标 BANF 前一行 `Verfügb. Menge` 中的可用数量。

Because MD04 normally already deducts safety stock from its running available quantity, safety stock remains protected in this calculation. The program does not add safety stock back to the usable quantity.

由于 MD04 通常已经在可用数量中扣除了安全库存，因此本计算会保留安全库存，不会把安全库存重新加回可用数量。

---

## Example | 示例

Available quantity immediately before the selected BANF:

目标 BANF 前的可用数量：

```text
Date:                 24.11.2026
Available quantity:   4
```

Selected BANF:

目标 BANF：

```text
Date:                 03.12.2026
BANF quantity:        30
```

Demand after the BANF until 31.12.2026:

目标 BANF 后至 31.12.2026 的需求：

```text
16 pcs
```

Other receipts after the BANF until 31.12.2026:

目标 BANF 后至 31.12.2026 的其他收货：

```text
0 pcs
```

Calculation:

计算：

```text
16 - 4 - 0 = 12
```

Result:

结果：

```text
BANF Quantity:       30
Required Quantity:   12
Excess Quantity:     18
```

中文说明：

```text
采购申请数量：        30
截止日期前真实需要：  12
截止日期前超额数量：  18
```

---

## Features | 功能

### Supported | 当前支持

- Read SAP MD04 Excel exports in `.xlsx` format
- 读取 SAP MD04 导出的 `.xlsx` 文件

- Select a specific BANF and item
- 支持指定 BANF 及行项目

- Enter a user-defined cutoff date
- 支持自定义截止日期

- Automatically select the latest BANF before the cutoff date
- 可自动选择截止日期前最近的一张 BANF

- Read the available quantity immediately before the selected BANF
- 读取目标 BANF 前一行的可用数量

- Calculate demand after the selected BANF until the cutoff date
- 统计目标 BANF 后至截止日期的需求

- Include other receipts after the selected BANF
- 考虑目标 BANF 后至截止日期的其他收货

- Calculate the required and excess BANF quantities
- 计算 BANF 真实需求量和超额数量

- Support German-language SAP MD04 exports
- 支持德语版 SAP MD04 导出文件

### Planned | 计划功能

- CSV result export
- 导出 CSV 结果

- Excel report generation
- 生成 Excel 分析报告

- Batch analysis of multiple materials
- 批量分析多个物料

- Graphical user interface
- 图形用户界面

- Configurable SAP language and MRP-element mappings
- 可配置的 SAP 语言及 MRP 元素映射

---

## Requirements | 环境要求

- Python 3.9 or later
- Python 3.9 或更高版本

- `openpyxl`
- SAP MD04 export in `.xlsx` format
- SAP MD04 导出的 `.xlsx` 文件

Install the dependency:

安装依赖：

```bash
python -m pip install openpyxl
```

---

## Project Structure | 项目结构

```text
sap-md04-banf-calculator/
├── md04_banf_realbedarf.py
├── README.md
├── requirements.txt
└── sample_data/
    └── sample_md04.xlsx
```

Only synthetic or fully anonymized sample data should be stored in a public repository.

公开仓库中只能保存合成数据或已完全匿名化的示例数据。

---

## Usage | 使用方法

### Interactive Mode | 交互模式

Run:

运行：

```bash
python md04_banf_realbedarf.py
```

Example input:

输入示例：

```text
Pfad zur MD04-Excel-Datei: EXPORT.xlsx
Stichtag (TT.MM.JJJJ): 31.12.2026
BANF/Position: 0000000001/00010
```

If the Excel file is in the same folder as the Python program, enter either:

如果 Excel 文件与 Python 程序位于同一个文件夹，可以输入：

```text
EXPORT.xlsx
```

or:

或者：

```text
.\EXPORT.xlsx
```

Do not leave the file path empty unless the program has been configured with a default file name.

如果程序没有设置默认文件名，请不要在文件路径处直接按 Enter。

### Command-Line Mode | 命令行模式

#### Windows PowerShell

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026 --banf 0000000001/00010
```

#### Linux or macOS

```bash
python md04_banf_realbedarf.py EXPORT.xlsx \
  --stichtag 31.12.2026 \
  --banf 0000000001/00010
```

The ISO date format is also supported:

也支持 ISO 日期格式：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 2026-12-31 --banf 0000000001/00010
```

### Automatic BANF Selection | 自动选择 BANF

If no BANF is provided, the program can select the most recent BANF whose planning date is on or before the cutoff date.

如果未指定 BANF，程序可以自动选择计划日期不晚于截止日期的最近一张 BANF。

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026
```

When prompted for the BANF, press Enter:

程序询问 BANF 时直接按 Enter：

```text
BANF/Position (Enter = letzte BANF bis Stichtag):
```

### Select a Worksheet | 指定工作表

The first worksheet is used by default. Use `--sheet` to select another worksheet:

程序默认读取第一个工作表。可以使用 `--sheet` 指定其他工作表：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --sheet Sheet1 --stichtag 31.12.2026 --banf 0000000001/00010
```

---

## Expected Excel Columns | Excel 列要求

The MD04 export should contain at least the following columns:

MD04 导出的 Excel 至少应包含以下列：

```text
Plantermine
Dispoel.
Daten zum DE
Zugang/Bedarf
Verfügb. Menge
```

Column meanings:

各列含义：

- `Plantermine`: Planning date | 计划日期
- `Dispoel.`: MRP element | MRP 元素
- `Daten zum DE`: MRP element data | MRP 元素数据
- `Zugang/Bedarf`: Receipt or requirement quantity | 收货或需求数量
- `Verfügb. Menge`: Running available quantity | 动态可用数量

The current implementation is designed primarily for German-language MD04 exports. Customized SAP systems may use different column names or MRP-element codes.

当前版本主要适用于德语版 MD04 导出文件。经过定制的 SAP 系统可能使用不同的列名或 MRP 元素代码。

---

## Recognized MD04 Elements | 支持的 MD04 元素

### Purchase Requisitions | 采购申请

```text
BS-Anf
BANF
PurRqs
```

### Common Demand Elements | 常见需求元素

```text
AR-Res
SekBed
Kund-B
Uml-Res
AB-Res
```

### Opening Stock | 初始库存

```text
BStand
```

### Safety Stock | 安全库存

```text
ShBest
```

Other negative `Zugang/Bedarf` values after the selected BANF may also be treated as demand, depending on the implementation of the script.

根据程序的具体实现，目标 BANF 后其他 `Zugang/Bedarf` 为负数的项目也可能被视为需求。

---

## Example Output | 输出示例

```text
=== Ergebnis ===
BANF:                              0000000001/00010
BANF-Termin:                       03.12.2026
BANF-Menge:                        30
Stichtag:                          31.12.2026
MD04-Bestand direkt vor BANF:      4
Bedarf nach BANF bis Stichtag:     16
Weitere Zugaenge nach BANF:        0
Saldo ohne Ziel-BANF:              -12
Benoetigte Menge aus Ziel-BANF:    12
Ueberdeckung der Ziel-BANF:        18
Restfehlmenge trotz voller BANF:   0
```

Interpretation:

结果说明：

```text
Original BANF quantity:            30
Required before the cutoff date:   12
Excess before the cutoff date:     18
Remaining shortage:                 0
```

对应中文：

```text
BANF 原始数量：                    30
截止日期前真实需要：                12
截止日期前超额数量：                18
完整保留 BANF 后仍存在的缺口：       0
```

---

## Output Fields | 输出字段

- `BANF`: Selected purchase requisition and item | 目标采购申请及行项目
- `BANF-Termin`: Planning date of the selected BANF | 目标 BANF 的计划日期
- `BANF-Menge`: Original BANF quantity | BANF 原始数量
- `Stichtag`: User-defined cutoff date | 用户指定的截止日期
- `MD04-Bestand direkt vor BANF`: Available quantity immediately before the target BANF | 目标 BANF 前一行的可用数量
- `Bedarf nach BANF bis Stichtag`: Demand after the target BANF until the cutoff date | 目标 BANF 后至截止日期的需求
- `Weitere Zugaenge nach BANF`: Other receipts after the target BANF until the cutoff date | 目标 BANF 后至截止日期的其他收货
- `Saldo ohne Ziel-BANF`: Projected balance without the selected BANF | 不考虑目标 BANF 时的预计余额
- `Benoetigte Menge aus Ziel-BANF`: Quantity required from the selected BANF | 目标 BANF 的真实需要量
- `Ueberdeckung der Ziel-BANF`: Quantity exceeding the calculated requirement | 超出真实需求的 BANF 数量
- `Restfehlmenge trotz voller BANF`: Remaining shortage after using the full BANF | 完整保留 BANF 后仍存在的缺口

---

## Assumptions | 计算假设

The tool assumes that:

本工具假设：

1. One Excel file contains the MD04 data for one material.
2. 一个 Excel 文件包含一个物料的 MD04 数据。

3. The Excel export is complete and sorted in the same sequence as MD04.
4. Excel 导出数据完整，并保持 MD04 中的行顺序。

5. The `Verfügb. Menge` value immediately before the target BANF is valid.
6. 目标 BANF 前一行的 `Verfügb. Menge` 有效。

7. Positive `Zugang/Bedarf` values are receipts.
8. `Zugang/Bedarf` 正数代表收货。

9. Negative `Zugang/Bedarf` values are requirements.
10. `Zugang/Bedarf` 负数代表需求。

11. The selected BANF is excluded from other receipts.
12. 目标 BANF 本身不会被计入其他收货。

13. Movements after the cutoff date are ignored.
14. 截止日期之后的需求和收货不会参与计算。

15. Other receipts are assumed to arrive on the dates shown in MD04.
16. 其他收货被假设能够在 MD04 显示的日期按时到达。

17. Safety stock remains protected because the MD04 available quantity is used directly.
18. 由于直接使用 MD04 的可用数量，因此安全库存保持不被占用。

---

## Limitations | 局限性

The tool does not determine whether:

本工具不会判断：

- A supplier will deliver on time | 供应商是否会按时交付
- A requirement will be postponed or cancelled | 需求是否会延期或取消
- A BANF has already been converted into a purchase order | BANF 是否已转为采购订单
- A supplier permits cancellation, postponement, or quantity reduction | 供应商是否允许取消、延期或减量
- Production schedules will change | 生产计划是否会发生变化
- Additional demand exists outside the exported MD04 data | SAP 导出数据之外是否还有额外需求
- SAP master data and MRP parameters are correct | SAP 主数据和 MRP 参数是否正确

The tool does not connect to SAP and does not modify any SAP object.

本工具不会连接 SAP，也不会修改任何 SAP 对象。

The result should be treated as analytical support and reviewed before purchasing or planning decisions are made.

计算结果仅用于分析支持。在做出采购或计划决策前，应进行人工复核。

---

## Troubleshooting | 常见问题

### Excel File Not Found | 找不到 Excel 文件

If the file is in the current folder, enter:

如果文件位于当前文件夹，请输入：

```text
EXPORT.xlsx
```

or:

或者：

```text
.\EXPORT.xlsx
```

You can check available Excel files in PowerShell with:

可以在 PowerShell 中检查当前目录下的 Excel 文件：

```powershell
Get-ChildItem *.xlsx
```

### Empty File Path | 文件路径为空

Do not press Enter without entering a file name unless the script defines a default workbook.

如果程序没有设定默认 Excel 文件，请不要在文件路径处直接按 Enter。

### Invalid Date | 日期无效

Use one of the supported formats without additional punctuation:

请使用以下日期格式，不要添加中文顿号或其他符号：

```text
31.12.2026
2026-12-31
```

Incorrect example:

错误示例：

```text
、31.12.2026
```

### BANF Not Found | 找不到 BANF

Check that:

请检查：

- The BANF and item number are complete | BANF 和行项目是否完整
- Leading zeros are preserved | 是否保留前导零
- The BANF exists in the exported Excel file | BANF 是否存在于导出的 Excel 中
- The MRP element is recognized by the script | 程序是否识别对应的 MRP 元素

### Unsupported Excel Format | 不支持的 Excel 格式

`openpyxl` supports modern Excel formats such as `.xlsx` and `.xlsm`, but it does not read the legacy `.xls` format.

`openpyxl` 支持 `.xlsx`、`.xlsm` 等现代 Excel 格式，但不能直接读取旧版 `.xls` 文件。

If necessary, open the file in Excel and save it as `.xlsx`.

如有需要，请使用 Excel 打开文件并另存为 `.xlsx`。

---

## Data Privacy | 数据隐私

Before uploading files to a public GitHub repository:

上传文件到公开 GitHub 仓库前：

- Remove or replace real material numbers | 删除或替换真实物料号
- Remove or replace purchase requisition numbers | 删除或替换真实采购申请号
- Remove purchase order numbers | 删除采购订单号
- Remove supplier and employee information | 删除供应商及员工信息
- Remove customer, plant, storage-location, and project identifiers | 删除客户、工厂、库存地点和项目标识
- Do not commit raw production exports | 不要提交真实生产环境导出文件
- Check the Git history if confidential data was committed previously | 如果曾提交敏感数据，请检查并清理 Git 历史

Recommended `.gitignore`:

推荐的 `.gitignore`：

```gitignore
# SAP and spreadsheet exports
*.xlsx
*.xls
*.csv

# Python cache
__pycache__/
*.py[cod]

# Virtual environments
.venv/
venv/

# Local editor settings
.vscode/
.idea/
```

If a sample workbook is published, use only synthetic or fully anonymized data and label it clearly as sample data.

如需上传示例工作簿，请仅使用合成数据或已完全匿名化的数据，并明确标注为示例数据。

---

## Disclaimer | 免责声明

This project is an independent analytical utility and is not affiliated with, endorsed by, or provided by SAP.

本项目是一个独立的分析工具，与 SAP 官方无任何关联，也未获得 SAP 官方背书或提供。

SAP and related product names are trademarks or registered trademarks of their respective owners.

SAP 及相关产品名称是其对应权利人的商标或注册商标。

The results depend on the completeness and accuracy of the exported data and the assumptions described above. Review all results before making purchasing, planning, production, inventory, or supplier decisions.

计算结果取决于导出数据的完整性、准确性以及上述计算假设。在做出采购、计划、生产、库存或供应商相关决策前，请复核所有结果。

---

## Contributing | 参与贡献

Issues and pull requests are welcome.

欢迎提交 Issue 和 Pull Request。

Useful contributions may include:

可贡献的功能包括：

- Support for additional SAP languages | 支持更多 SAP 语言
- Configurable MRP-element mappings | 可配置的 MRP 元素映射
- Batch processing for multiple materials | 多物料批量处理
- Export of results to Excel or CSV | 将结果导出为 Excel 或 CSV
- Automated tests using synthetic MD04 data | 使用合成 MD04 数据的自动化测试
- A graphical user interface | 图形用户界面

When contributing examples or test files, use only synthetic or fully anonymized data.

提交示例或测试文件时，只能使用合成数据或已完全匿名化的数据。

---

## License | 许可证

This project can be published under the MIT License. Add a `LICENSE` file to the repository before distribution.

本项目可以使用 MIT License 发布。正式分发前，请在仓库中添加 `LICENSE` 文件。
