using System;
using ClosedXML.Excel;

namespace ExportIncentiveApp.Services
{
    public static class TemplateService
    {
        public static void CreateTemplate(string savePath)
        {
            using var wb = new XLWorkbook();

            // Create 2 sample sheets
            for (int s = 1; s <= 2; s++)
            {
                var ws = wb.Worksheets.Add($"Sheet{s}");

                string[] headers = {
                    "company_name", "incentive_type", "circular_no", "auditor_reference",
                    "bank_name",
                    "contract_number", "contract_date", "contract_value",  // ← contract info
                    "exchange_rate", "incentive_percentage", "application_date",
                    "exp_number", "export_value_usd", "repatriated_amount_usd",
                    "freight_charge", "commission_insurance", "destination_country",
                    "shipment_date", "repatriation_date"
                };

                // Header row
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(1, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    cell.Style.Alignment.WrapText = true;
                }

                string companyName  = s == 1 ? "ABC Garments Ltd" : "XYZ Export Ltd";
                string incType      = s == 1 ? "Special Export Subsidy" : "Cash Incentive";
                string circular     = "FE Circular 12/2024";
                string audRef       = $"MS/2025/0{s}";
                string bank         = "United Bank Ltd";
                string contractNo   = $"LA{3020 + s}";
                string contractDate = "15/05/2024";
                double contractVal  = 50000;
                double exRate       = s == 1 ? 110.0 : 121.0;
                double incPct       = s == 1 ? 20.0 : 5.0;
                string appDate      = "15/10/2025";

                // Two data rows per sheet
                for (int r = 0; r < 2; r++)
                {
                    int row = r + 2;
                    ws.Cell(row, 1).Value  = companyName;
                    ws.Cell(row, 2).Value  = incType;
                    ws.Cell(row, 3).Value  = circular;
                    ws.Cell(row, 4).Value  = audRef;
                    ws.Cell(row, 5).Value  = bank;
                    ws.Cell(row, 6).Value  = contractNo;
                    ws.Cell(row, 7).Value  = contractDate;
                    ws.Cell(row, 8).Value  = contractVal;
                    ws.Cell(row, 9).Value  = exRate;
                    ws.Cell(row, 10).Value = incPct;
                    ws.Cell(row, 11).Value = appDate;
                    ws.Cell(row, 12).Value = $"EXP-00{r + 1}";
                    ws.Cell(row, 13).Value = 50000;
                    ws.Cell(row, 14).Value = 48000;
                    ws.Cell(row, 15).Value = 1000;
                    ws.Cell(row, 16).Value = 500;
                    ws.Cell(row, 17).Value = "Italy";
                    ws.Cell(row, 18).Value = r == 0 ? "15/10/2024" : "16/10/2024";
                    ws.Cell(row, 19).Value = r == 0 ? "01/12/2024" : "02/12/2024";
                }

                ws.Columns().AdjustToContents();
            }

            wb.SaveAs(savePath);
        }
    }
}
