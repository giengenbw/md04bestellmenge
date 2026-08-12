# SAP MD04 BANF Real Demand Calculator

A lightweight Python tool that reads an Excel export from SAP MD04 and estimates how much of a selected purchase requisition, known as a **BANF** in German SAP terminology, is actually required before a user-defined cutoff date.

The calculator uses the following business rule:

> Calculate actual production demand while allowing safety stock to be consumed.

For this reason, the MD04 safety stock element `ShBest` is shown for reference but is excluded from the net real-demand calculation.

## Features

- Reads SAP MD04 exports in `.xlsx` format
- Lets the user enter a cutoff date
- Supports selection of a specific purchase requisition and item
- Can automatically select the latest purchase requisition before the cutoff date
- Separates safety stock from actual demand
- Includes other receipts before the cutoff date
- Reports:
  - Actual quantity required from the selected purchase requisition
  - Excess purchase requisition quantity
  - Remaining shortage if the full requisition is insufficient
  - Opening stock, total actual demand, other receipts, and safety stock

## Calculation Method

### Actual net requirement

```text
Actual net requirement
= Actual demand up to the cutoff date
- Opening stock
- Other valid receipts
```

### Quantity required from the selected purchase requisition

```text
Required purchase requisition quantity
= min(
    Original purchase requisition quantity,
    max(0, Actual net requirement)
  )
```

### Excess quantity

```text
Excess quantity
= Original purchase requisition quantity
- Required purchase requisition quantity
```

## Safety Stock Treatment

The tool intentionally excludes `ShBest` from actual demand.

This means safety stock is assumed to be available for consumption. The result therefore represents the shortage caused by actual demand, not the replenishment quantity required to preserve the safety-stock level.

If safety stock must remain untouched, this calculation method is not suitable without modification.

## Example

Assume the MD04 data contains:

```text
Opening stock:                         20
Actual demand up to the cutoff date:   31
Other valid receipts:                   0
Safety stock:                           3  (information only)
Purchase requisition quantity:         50
```

The actual net requirement is:

```text
31 - 20 - 0 = 11
```

The result is therefore:

```text
Required purchase requisition quantity: 11
Excess purchase requisition quantity:   39
```

## Requirements

- Python 3.9 or later
- `openpyxl`
- An SAP MD04 export in `.xlsx` format

Install the dependency:

```bash
python -m pip install openpyxl
```

## Project Structure

```text
sap-md04-banf-calculator/
├── md04_banf_realbedarf.py
├── README.md
└── sample_md04.xlsx          # optional anonymized example file
```

Do not upload real SAP exports containing material numbers, purchase requisition numbers, supplier data, employee data, or other confidential information. Use anonymized sample data for public repositories.

## Usage

### Interactive mode

```bash
python md04_banf_realbedarf.py
```

The program asks for:

```text
Pfad zur MD04-Excel-Datei:
Stichtag (TT.MM.JJJJ):
BANF/Position (Enter = letzte BANF bis Stichtag):
```

Example input:

```text
Pfad zur MD04-Excel-Datei: sample_md04.xlsx
Stichtag (TT.MM.JJJJ): 31.12.2026
BANF/Position: 0000000001/00010
```

### Command-line mode

```bash
python md04_banf_realbedarf.py sample_md04.xlsx \
  --stichtag 31.12.2026 \
  --banf 0000000001/00010
```

The ISO date format is also supported:

```bash
python md04_banf_realbedarf.py sample_md04.xlsx \
  --stichtag 2026-12-31 \
  --banf 0000000001/00010
```

### Automatic purchase requisition selection

Omit `--banf` to select the latest purchase requisition whose planning date is on or before the cutoff date:

```bash
python md04_banf_realbedarf.py sample_md04.xlsx \
  --stichtag 31.12.2026
```

When prompted for the purchase requisition, press Enter.

### Select a worksheet

The first worksheet is used by default. To select another worksheet:

```bash
python md04_banf_realbedarf.py sample_md04.xlsx \
  --sheet Sheet1 \
  --stichtag 31.12.2026 \
  --banf 0000000001/00010
```

## Expected Excel Columns

The program identifies the required fields by their SAP MD04 column headings. The workbook should contain at least:

```text
Plantermine
Dispoel.
Daten zum DE
Zugang/Bedarf
Verfügb. Menge
```

Their expected meanings are:

- `Plantermine`: Planning date
- `Dispoel.`: MRP element
- `Daten zum DE`: MRP element data, such as the purchase requisition and item number
- `Zugang/Bedarf`: Receipt as a positive quantity or requirement as a negative quantity
- `Verfügb. Menge`: Running available quantity calculated by MD04

The current implementation is designed for German-language MD04 exports. Column aliases or MRP-element mappings may need to be added for exports in other languages or customized SAP environments.

## Recognized MRP Elements

The script currently recognizes the following element groups:

### Purchase requisitions

```text
BS-Anf
BANF
PurRqs
```

### Demand elements

```text
AR-Res
SekBed
Kund-B
Uml-Res
AB-Res
```

### Opening stock

```text
BStand
```

### Safety stock

```text
ShBest
```

If an export uses different MRP-element codes, update the corresponding sets near the beginning of the Python script.

## Example Output

```text
=== Ergebnis ===
BANF:                         0000000001/00010
BANF-Termin:                  23.10.2026
BANF-Menge:                   50
Stichtag:                     31.12.2026
Saldo ohne Ziel-BANF:         -11
Realer BANF-Bedarf:           11
Ueberdeckung der BANF:        39
Fehlmenge trotz voller BANF:  0

Kontrollwerte:
  Anfangsbestand:             20
  Sonstige Zugaenge:          0
  Realer Bedarf:              31
  Sicherheitsbestand (Info):  3
  Bruttobedarf nach BANF:     15
```

## Output Fields

- `BANF`: Selected purchase requisition and item
- `BANF-Termin`: Planning date of the selected requisition
- `BANF-Menge`: Original requisition quantity
- `Stichtag`: User-defined cutoff date
- `Saldo ohne Ziel-BANF`: Projected balance without the selected requisition and without treating safety stock as demand
- `Realer BANF-Bedarf`: Quantity actually required from the selected requisition
- `Ueberdeckung der BANF`: Quantity exceeding the calculated real requirement
- `Fehlmenge trotz voller BANF`: Remaining shortage after using the complete selected requisition
- `Anfangsbestand`: Opening physical stock read from the `BStand` row
- `Sonstige Zugaenge`: Positive receipts other than the selected requisition
- `Realer Bedarf`: Negative demand elements up to the cutoff date, excluding `ShBest`
- `Sicherheitsbestand (Info)`: Safety stock shown for information only
- `Bruttobedarf nach BANF`: Actual demand after the selected requisition row and up to the cutoff date

## Assumptions

The current calculation assumes that:

1. One workbook contains the MD04 detail for one material.
2. The `Verfügb. Menge` value on the relevant `BStand` row represents opening physical stock.
3. Positive `Zugang/Bedarf` values are receipts.
4. Negative `Zugang/Bedarf` values are requirements.
5. `ShBest` is excluded from actual demand.
6. The selected purchase requisition is excluded from other receipts.
7. Movements after the cutoff date are ignored.
8. Outstanding demand displayed after the `BStand` row remains relevant even when its planning date is earlier than the current date.
9. Other receipts are assumed to arrive on the dates shown in MD04.
10. The Excel export is complete and reflects the latest available SAP data.

## Limitations

The tool does not determine whether:

- A receipt will arrive on time
- A requirement will be postponed or cancelled
- A purchase requisition has already been converted into a purchase order
- A supplier permits cancellation, postponement, or quantity reduction
- Safety stock may be consumed under the applicable business policy
- Additional demand exists outside the exported MD04 data

The tool does not connect to SAP and does not modify any SAP object.

## Troubleshooting

### File not found

Use the complete path to the workbook. Put paths containing spaces in quotation marks:

```bash
python md04_banf_realbedarf.py "C:\path with spaces\sample_md04.xlsx" \
  --stichtag 31.12.2026
```

### Purchase requisition not found

Check that:

- The requisition and item number are complete
- Leading zeros are preserved
- The requisition exists in the exported workbook
- The corresponding MRP element is recognized by the script

### Opening stock not found

The script requires a `BStand` row before the selected purchase requisition. Export a complete MD04 list if that row is missing.

### Results differ from SAP available quantity

This can be expected. The tool excludes safety stock from real demand, while MD04 may deduct safety stock when calculating its running available quantity.

## Data Privacy

Before publishing files to a public repository:

- Remove or replace real material numbers
- Remove or replace purchase requisition and purchase order numbers
- Remove supplier and employee information
- Remove plant, storage-location, project, and customer identifiers where necessary
- Do not commit raw production exports
- Check the Git history if confidential data was committed previously

A `.gitignore` entry can help prevent accidental uploads:

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
```

If you want to publish an example workbook, create a fully anonymized or synthetic file and explicitly label it as sample data.

## Disclaimer

This project is an independent utility for analytical support. It is not affiliated with, endorsed by, or provided by SAP. SAP and related product names are trademarks or registered trademarks of their respective owners.

The results depend on the completeness and accuracy of the exported data and the assumptions described above. Review the result before making purchasing, planning, production, inventory, or supplier decisions.

## Contributing

Issues and pull requests are welcome. Useful contributions may include:

- Support for additional SAP languages
- Configurable MRP-element mappings
- Batch processing for multiple materials
- Export of results to Excel or CSV
- Automated tests with synthetic MD04 data
- A graphical user interface

When contributing examples or test files, use only synthetic or fully anonymized data.

## License

No license is included by default. Before publishing the repository, add a license that matches your intended use, such as the MIT License.
