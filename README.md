# SAP MD04 BANF 真实需求量计算程序

本程序用于读取 SAP MD04 导出的 Excel 文件，并根据用户输入的截止日期，计算指定采购申请（BANF）在截止日期前的真实需求量。

本程序采用以下计算口径：

> **允许使用安全库存。**

因此，MD04 中的安全库存项目 `ShBest` 只作为参考信息显示，不作为实际需求参与计算。换句话说，安全库存可以用于满足截止日期前的生产需求。

---

## 一、程序能做什么

程序可以完成以下操作：

- 读取 SAP MD04 导出的 `.xlsx` 文件
- 手动输入计算截止日期
- 手动指定采购申请号和行项目
- 未指定 BANF 时，自动选择截止日期前最近的一张 BANF
- 读取 MD04 中的初始库存
- 汇总截止日期前的实际需求
- 汇总除目标 BANF 以外的其他入库
- 将安全库存排除在实际需求之外
- 计算目标 BANF 的真实需要量
- 计算目标 BANF 的超额数量
- 判断完整保留目标 BANF 后是否仍有缺口

---

## 二、计算口径

### 1. 允许使用安全库存

程序会识别 MD04 中的安全库存项目：

```text
ShBest
```

但是 `ShBest` 不计入实际需求。

例如，MD04 中存在：

```text
当前库存：      50
安全库存：      10
实际需求：     112
其他入库：      60
目标 BANF：     30
```

程序计算真实净需求时使用：

```text
真实净需求
= 实际需求 - 当前库存 - 其他入库
= 112 - 50 - 60
= 2
```

安全库存 10 不会再额外增加采购需求，因为本程序允许使用安全库存。

### 2. 核心公式

```text
真实净需求
= 截止日期前的实际需求
- 当前实际库存
- 截止日期前除目标 BANF 外的其他入库
```

目标 BANF 的真实需要量：

```text
目标 BANF 真实需要量
= 原 BANF 数量与真实净需求中的较小值
```

同时，真实需要量不会小于 0：

```text
目标 BANF 真实需要量
= min(
    BANF 原始数量,
    max(0, 真实净需求)
  )
```

BANF 超额数量：

```text
BANF 超额数量
= BANF 原始数量 - BANF 真实需要量
```

如果真实净需求大于 BANF 原始数量，程序还会显示完整保留 BANF 后仍然存在的缺口。

---

## 三、运行环境

### 必要条件

- Windows、Linux 或 macOS
- Python 3.9 或更高版本
- `openpyxl` 模块
- SAP MD04 导出的 `.xlsx` 文件

建议使用 Python 3.10 或更高版本。

### 安装 Python 模块

打开 PowerShell、命令提示符或终端，执行：

```powershell
python -m pip install openpyxl
```

如果电脑上使用的是 `python3` 命令，可以执行：

```powershell
python3 -m pip install openpyxl
```

检查是否安装成功：

```powershell
python -c "import openpyxl; print(openpyxl.__version__)"
```

---

## 四、准备文件

建议把 Python 程序和 MD04 Excel 文件放在同一个文件夹中。

示例：

```text
C:\md04bestellmenge\
├── md04_banf_realbedarf.py
├── EXPORT.xlsx
└── README.md
```

其中：

- `md04_banf_realbedarf.py` 是计算程序
- `EXPORT.xlsx` 是从 SAP MD04 导出的 Excel 文件
- `README.md` 是使用说明

### 当前文件夹的表示方法

在 Windows PowerShell 中：

```text
.
```

表示当前文件夹。

因此，以下两种输入含义相同：

```text
EXPORT.xlsx
```

```text
.\EXPORT.xlsx
```

如果 PowerShell 当前位于：

```text
C:\md04bestellmenge
```

那么：

```text
.\EXPORT.xlsx
```

代表：

```text
C:\md04bestellmenge\EXPORT.xlsx
```

---

## 五、确认 Excel 文件存在

在运行程序前，建议先在 PowerShell 中执行：

```powershell
Get-ChildItem *.xlsx
```

也可以使用：

```powershell
dir *.xlsx
```

如果文件存在，应当能看到类似结果：

```text
EXPORT.xlsx
```

如果没有看到文件，请检查：

- Excel 文件是否在当前文件夹
- 文件名是否正确
- 文件扩展名是否为 `.xlsx`
- 文件是否被保存成旧版 `.xls`

---

## 六、使用方法

程序支持两种运行方式：

1. 交互式运行
2. 命令行参数运行

推荐使用命令行参数运行，因为输入更快，也可以避免忘记填写 Excel 文件路径。

---

## 七、交互式运行

在 PowerShell 中进入程序所在文件夹：

```powershell
cd C:\md04bestellmenge
```

然后运行：

```powershell
python md04_banf_realbedarf.py
```

程序会依次询问：

```text
Pfad zur MD04-Excel-Datei:
Stichtag (TT.MM.JJJJ):
BANF/Position (Enter = letzte BANF bis Stichtag):
```

输入示例：

```text
Pfad zur MD04-Excel-Datei: EXPORT.xlsx
Stichtag (TT.MM.JJJJ): 31.12.2026
BANF/Position (Enter = letzte BANF bis Stichtag): 0000000001/00010
```

### Excel 文件位于当前文件夹

输入：

```text
EXPORT.xlsx
```

或者：

```text
.\EXPORT.xlsx
```

### Excel 文件位于其他文件夹

请输入完整路径，例如：

```text
C:\Users\用户名\Downloads\EXPORT.xlsx
```

如果路径中包含空格，也可以直接输入完整路径。交互模式下通常不需要添加引号。

### 注意：不要把文件路径留空

出现以下提示时：

```text
Pfad zur MD04-Excel-Datei:
```

必须输入 Excel 文件名或完整路径。

如果直接按 Enter，程序可能会把当前文件夹当成 Excel 文件，从而报错：

```text
InvalidFileException
```

正确输入示例：

```text
EXPORT.xlsx
```

---

## 八、命令行参数运行

### 指定 Excel、截止日期和 BANF

Windows PowerShell：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026 --banf 0000000001/00010
```

参数含义：

```text
EXPORT.xlsx              MD04 Excel 文件
--stichtag 31.12.2026    截止日期
--banf 0000000001/00010  采购申请号及行项目
```

### 使用完整 Excel 路径

```powershell
python md04_banf_realbedarf.py "C:\SAP Export\EXPORT.xlsx" --stichtag 31.12.2026 --banf 0000000001/00010
```

路径中有空格时，必须使用双引号。

### 使用 ISO 日期格式

除 `31.12.2026` 外，也可以使用：

```text
2026-12-31
```

运行示例：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 2026-12-31 --banf 0000000001/00010
```

---

## 九、自动选择 BANF

如果不输入 `--banf`，程序会要求手动输入 BANF：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026
```

程序显示：

```text
BANF/Position (Enter = letzte BANF bis Stichtag):
```

此时直接按 Enter，程序会自动选择：

```text
计划日期不晚于截止日期的最近一张 BANF
```

如果截止日期前存在多张 BANF，程序按照日期和 Excel 中的行位置选择最近的一张。

为了避免选错 BANF，实际使用时建议明确输入 BANF 和行项目。

---

## 十、指定 Excel 工作表

程序默认读取 Excel 的第一个工作表。

如果 MD04 数据位于指定工作表，例如 `Sheet1`，可以使用：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --sheet Sheet1 --stichtag 31.12.2026 --banf 0000000001/00010
```

如果工作表名称包含空格，需要使用双引号：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --sheet "MD04 Export" --stichtag 31.12.2026 --banf 0000000001/00010
```

---

## 十一、Excel 文件要求

程序至少需要以下列：

```text
Plantermine
Dispoel.
Daten zum DE
Zugang/Bedarf
Verfügb. Menge
```

各列含义：

- `Plantermine`：计划日期
- `Dispoel.`：MRP 元素
- `Daten zum DE`：MRP 元素数据，例如 BANF 和行项目
- `Zugang/Bedarf`：收货或需求数量
- `Verfügb. Menge`：MD04 动态可用数量

程序主要适用于德语版 SAP MD04 导出文件。

如果 SAP 系统中的列名经过定制，可能需要修改程序中的列名候选项。

---

## 十二、程序识别的 MD04 元素

### 采购申请

```text
BS-Anf
BANF
PurRqs
```

### 常见需求元素

```text
AR-Res
SekBed
Kund-B
Uml-Res
AB-Res
```

### 初始库存

```text
BStand
```

### 安全库存

```text
ShBest
```

安全库存 `ShBest` 只作为参考值显示，不参与实际需求合计。

---

## 十三、输出示例

```text
=== Ergebnis ===
BANF:                         0000000001/00010
BANF-Termin:                  03.12.2026
BANF-Menge:                   30
Stichtag:                     31.12.2026
Saldo ohne Ziel-BANF:         -2
Realer BANF-Bedarf:           2
Ueberdeckung der BANF:        28
Fehlmenge trotz voller BANF:  0

Kontrollwerte:
  Anfangsbestand:             50
  Sonstige Zugaenge:          60
  Realer Bedarf:              112
  Sicherheitsbestand (Info):  10
  Bruttobedarf nach BANF:     16
```

---

## 十四、输出字段说明

### `BANF`

目标采购申请号和行项目。

### `BANF-Termin`

目标 BANF 在 MD04 中的计划日期。

### `BANF-Menge`

目标 BANF 的原始数量。

### `Stichtag`

用户输入的截止日期。

### `Saldo ohne Ziel-BANF`

不考虑目标 BANF 时，截止日期前满足实际需求后的预计库存余额。

如果结果为负数，负数的绝对值表示需要由目标 BANF 补足的数量。

例如：

```text
Saldo ohne Ziel-BANF: -2
```

表示不考虑目标 BANF 时缺少 2 个。

### `Realer BANF-Bedarf`

目标 BANF 截止该日期真正需要的数量。

例如：

```text
Realer BANF-Bedarf: 2
```

表示目标 BANF 中有 2 个在截止日期前真正需要。

### `Ueberdeckung der BANF`

目标 BANF 超过真实净需求的数量。

例如：

```text
BANF-Menge:            30
Realer BANF-Bedarf:     2
Ueberdeckung:          28
```

表示截止日期前暂时不需要其中的 28 个。

### `Fehlmenge trotz voller BANF`

即使完整保留目标 BANF，仍然存在的缺口。

如果为 0，表示目标 BANF 足以覆盖真实净需求。

### `Anfangsbestand`

从 `BStand` 行的 `Verfügb. Menge` 读取的初始库存。

### `Sonstige Zugaenge`

截止日期前除目标 BANF 以外的其他正数入库。

这些入库可能包括其他采购申请、采购订单、计划订单或其他 MD04 收货元素。

### `Realer Bedarf`

截止日期前所有实际负数需求的合计，不包含安全库存 `ShBest`。

### `Sicherheitsbestand (Info)`

MD04 中的安全库存数量，只作为参考显示。

因为本程序允许使用安全库存，所以安全库存不会增加目标 BANF 的真实需求量。

### `Bruttobedarf nach BANF`

目标 BANF 行之后至截止日期的实际需求总量，不包含安全库存。

该数值用于检查目标 BANF 之后还有多少需求，但目标 BANF 的真实需要量仍由完整的净需求公式计算。

---

## 十五、完整使用示例

假设程序和 Excel 文件均位于：

```text
C:\md04bestellmenge
```

首先进入该文件夹：

```powershell
cd C:\md04bestellmenge
```

检查文件：

```powershell
Get-ChildItem
```

确认存在：

```text
md04_banf_realbedarf.py
EXPORT.xlsx
```

安装依赖：

```powershell
python -m pip install openpyxl
```

运行程序：

```powershell
python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026 --banf 0000000001/00010
```

如果 PowerShell 中的 `python` 命令不可用，可以使用完整 Python 路径：

```powershell
& "C:\Path\To\python.exe" .\md04_banf_realbedarf.py .\EXPORT.xlsx --stichtag 31.12.2026 --banf 0000000001/00010
```

---

## 十六、常见错误及解决方法

### 1. `from __future__ imports must occur at the beginning of the file`

错误原因：

`from __future__ import annotations` 前面出现了普通 Python 语句。

错误示例：

```python
'python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026'
from __future__ import annotations
```

正确写法是把运行命令改成注释：

```python
# 运行示例：
# python md04_banf_realbedarf.py EXPORT.xlsx --stichtag 31.12.2026

from __future__ import annotations
```

也可以删除 `from __future__ import annotations`。在 Python 3.10 或更高版本中，本程序使用的类型标注通常不依赖该语句。

### 2. `Datum ungueltig`

错误原因：

日期格式不正确，或者日期前后含有其他符号。

正确输入：

```text
31.12.2026
```

或者：

```text
2026-12-31
```

错误输入：

```text
、31.12.2026
```

请删除日期前面的中文顿号或其他符号。

### 3. `InvalidFileException`

常见原因：

- Excel 文件路径为空
- 输入的是文件夹而不是 Excel 文件
- 文件格式不是 `.xlsx`
- 文件扩展名错误

正确输入：

```text
EXPORT.xlsx
```

或者：

```text
.\EXPORT.xlsx
```

检查文件是否存在：

```powershell
Get-ChildItem *.xlsx
```

### 4. 找不到 BANF

请检查：

- BANF 号码是否完整
- 行项目是否完整
- 是否保留了前导零
- BANF 是否存在于 Excel 中
- MRP 元素是否为程序支持的类型

正确示例：

```text
0000000001/00010
```

不要删除采购申请号或行项目中的前导零。

### 5. 找不到 `BStand`

程序需要从 `BStand` 行读取初始库存。

如果 Excel 中没有 `BStand`：

- 检查导出的 MD04 数据是否完整
- 重新从 SAP MD04 导出完整列表
- 确认没有手动删除 Excel 顶部数据

### 6. `.xls` 文件无法读取

`openpyxl` 不支持旧版 `.xls` 文件。

请使用 Excel 打开文件，然后选择：

```text
另存为 -> Excel 工作簿 (*.xlsx)
```

再使用程序读取新的 `.xlsx` 文件。

---

## 十七、使用结果时的注意事项

本程序只根据 Excel 中现有的 MD04 数据进行计算，不会判断：

- 其他收货是否会按时到达
- 供应商是否延迟交付
- 需求是否会取消或延期
- BANF 是否已经转成采购订单
- 采购订单是否已经确认
- 是否存在尚未录入 SAP 的额外需求
- 安全库存是否根据实际业务允许被使用
- SAP 主数据和 MRP 参数是否正确

在联系供应商减量、延期或取消之前，建议结合以下信息复核：

- 最新 MD04
- 采购订单状态
- 供应商确认日期
- 生产计划
- 库存盘点结果
- 计划部门意见

---

## 十八、数据隐私

不要把真实 SAP 导出文件直接上传到公开 GitHub 仓库。

公开代码前，请删除或匿名化：

- 真实物料号
- 真实 BANF 号码
- 采购订单号
- 供应商名称和编号
- 客户信息
- 员工信息
- 工厂和库存地点
- 项目编号
- 公司内部备注

推荐使用以下 `.gitignore`：

```gitignore
# SAP 和 Excel 导出文件
*.xlsx
*.xls
*.csv

# Python 缓存
__pycache__/
*.py[cod]

# 虚拟环境
.venv/
venv/

# 编辑器配置
.vscode/
.idea/
```

如果需要公开示例 Excel，请使用合成数据或完全匿名化的数据。

---

## 十九、免责声明

本程序是独立开发的分析辅助工具，与 SAP 官方无关，也未获得 SAP 官方背书。

SAP 及相关产品名称是其对应权利人的商标或注册商标。

计算结果取决于 MD04 导出数据的完整性和准确性。程序不会自动连接 SAP，也不会修改 SAP 中的采购申请、采购订单、库存或需求。

在做出采购、计划、生产、库存或供应商相关决策前，请对计算结果进行人工复核。

---

## 二十、许可证

如需将项目发布到 GitHub，可以添加 MIT License。

建议在项目根目录创建：

```text
LICENSE
```

并在正式发布前确认许可证符合实际使用需求。
