# Export Incentive PDF Generator — .NET WPF App

## Requirements
- Windows 10/11
- .NET 10 SDK (already installed ✓)
- VS Code with C# Dev Kit extension

---

## Project Structure

```
ExportIncentiveApp/
├── App.xaml
├── App.xaml.cs
├── MainWindow.xaml          ← UI layout
├── MainWindow.xaml.cs       ← UI logic
├── ExportIncentiveApp.csproj
├── Models/
│   └── Models.cs            ← Data models
└── Services/
    ├── HtmlTemplateService.cs  ← HTML templates (Certificate, Annexure, Bill)
    ├── PdfGeneratorService.cs  ← PDF generation logic
    └── TemplateService.cs      ← Excel template creator
```

---

## Step-by-Step Setup

### 1. Open in VS Code
```
cd ExportIncentiveApp
code .
```

### 2. Restore NuGet packages
```
dotnet restore
```

### 3. Download DinkToPdf native library (REQUIRED)
DinkToPdf needs `libwkhtmltox.dll` (64-bit) in your project folder.

Download from:
https://github.com/rdvojmoc/DinkToPdf/raw/master/v0.12.4/64%20bit/libwkhtmltox.dll

Place it in the root of `ExportIncentiveApp/` folder.

### 4. Build and Run
```
dotnet build
dotnet run
```

Or press **F5** in VS Code (install C# Dev Kit extension first).

---

## How to Use the App

1. **Browse Excel File** — select your `.xlsx` file
   - Each sheet = one contract
   - Use "Download Excel Template" to get the correct format

2. **Edit Contract Details** — Contract No., Date, Value are editable in the table

3. **Fill General Info** — Company Address, Bank Branch, Audit Fee Note

4. **Click "Generate PDF Documents"** — choose output folder
   - Generates: Certificate + Annexure PDF per contract
   - Generates: Audit Details Report (Bill) PDF
   - All files zipped into `Export_Documents_YYYYMMDD_HHMMSS.zip`

---

## Excel File Format

Each sheet must have these columns (row 1 = headers, row 2+ = data):

| Column | Description |
|--------|-------------|
| company_name | Exporter company name |
| incentive_type | e.g. "Cash Incentive" |
| circular_no | e.g. "Circular 01/2024" |
| auditor_reference | Auditor ref number |
| bank_name | Bank name |
| exchange_rate | BDT per USD |
| incentive_percentage | e.g. 10 (for 10%) |
| application_date | dd/MM/yyyy |
| exp_number | EXP number |
| export_value_usd | Export invoice value |
| repatriated_amount_usd | Repatriated amount |
| freight_charge | Freight deduction |
| commission_insurance | Commission/insurance |
| destination_country | e.g. Italy |
| shipment_date | dd/MM/yyyy |
| repatriation_date | dd/MM/yyyy |
| contract_value | LC/Contract USD value |

---

## Audit Fee Calculation
| Incentive (BDT) | Audit Fee |
|-----------------|-----------|
| 0 – 5,00,000 | 4,000 Taka |
| 5,00,001 – 10,00,000 | 5,000 Taka |
| 10,00,001+ | 7,000 Taka |
