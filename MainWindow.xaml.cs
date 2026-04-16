using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExportIncentiveApp.Models;
using ExportIncentiveApp.Services;

using MessageBox = System.Windows.MessageBox;
using MessageBoxButton = System.Windows.MessageBoxButton;
using MessageBoxImage = System.Windows.MessageBoxImage;
using MessageBoxResult = System.Windows.MessageBoxResult;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace ExportIncentiveApp
{
    public partial class MainWindow : System.Windows.Window
    {
        private List<ContractInfo> _contractsData = new();
        private ObservableCollection<ContractRow> _contractRows = new();

        public MainWindow()
        {
            InitializeComponent();
            ContractTable.ItemsSource = _contractRows;
        }

        private void BrowseFile_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                Title = "Select Excel File"
            };
            if (dlg.ShowDialog() == true)
            {
                FilePathLabel.Text = Path.GetFileName(dlg.FileName);
                FilePathLabel.Foreground = System.Windows.Media.Brushes.Black;
                LoadExcelData(dlg.FileName);
            }
        }

        private void LoadExcelData(string filePath)
        {
            try
            {
                _contractsData.Clear();
                _contractRows.Clear();

                using var wb = new XLWorkbook(filePath);

                foreach (var ws in wb.Worksheets)
                {
                    string sheetName = ws.Name;
                    string status;
                    string companyName = sheetName;

                    try
                    {
                        var usedRows = ws.RowsUsed().ToList();
                        if (usedRows.Count <= 1)
                        {
                            status = "Empty sheet";
                            _contractRows.Add(new ContractRow
                            {
                                SheetName = sheetName, CompanyName = companyName,
                                ContractNumber = "-", ContractDate = "-",
                                ContractValue = "-", Status = status
                            });
                            continue;
                        }

                        // Build column map from header row
                        var colMap = new Dictionary<string, int>();
                        var headerRow = ws.Row(1);
                        for (int c = 1; c <= (headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0); c++)
                        {
                            var val = headerRow.Cell(c).GetString().Trim().ToLower();
                            if (!string.IsNullOrEmpty(val)) colMap[val] = c;
                        }

                        int Col(string name) => colMap.TryGetValue(name, out int idx) ? idx : -1;

                        var firstRow = ws.Row(2);

                        // Read ALL contract info directly from Excel columns
                        companyName        = GetStr(firstRow, Col("company_name"), sheetName);
                        string contractNo  = GetStr(firstRow, Col("contract_number"), $"CT-{DateTime.Now.Year}-001");
                        string contractDt  = FormatDateCell(ws, 2, Col("contract_date"));
                        double contractVal = GetDouble(firstRow, Col("contract_value"), 0);
                        double exRate      = GetDouble(firstRow, Col("exchange_rate"), 1);
                        double incPct      = GetDouble(firstRow, Col("incentive_percentage"), 0);
                        string incType     = GetStr(firstRow, Col("incentive_type"), "Cash Incentive");
                        string circNo      = GetStr(firstRow, Col("circular_no"), "");
                        string audRef      = GetStr(firstRow, Col("auditor_reference"), "");
                        string bankName    = GetStr(firstRow, Col("bank_name"), "");
                        string appDate     = FormatDateCell(ws, 2, Col("application_date"));

                        double totalExport = 0, totalRepat = 0, totalFreight = 0, totalComm = 0;
                        var exports = new List<ExportRecord>();
                        var shipDates = new List<string>();

                        foreach (var row in usedRows.Skip(1))
                        {
                            double expVal  = GetDouble(row, Col("export_value_usd"), 0);
                            double repat   = GetDouble(row, Col("repatriated_amount_usd"), 0);
                            double freight = GetDouble(row, Col("freight_charge"), 0);
                            double comm    = GetDouble(row, Col("commission_insurance"), 0);
                            string shipD   = FormatDateCell(ws, row.RowNumber(), Col("shipment_date"));
                            string repatD  = FormatDateCell(ws, row.RowNumber(), Col("repatriation_date"));

                            totalExport  += expVal;
                            totalRepat   += repat;
                            totalFreight += freight;
                            totalComm    += comm;
                            if (!string.IsNullOrEmpty(shipD)) shipDates.Add(shipD);

                            double rowNetFob = repat - (freight + comm);
                            double rowInc    = rowNetFob * (incPct / 100.0) * exRate;

                            exports.Add(new ExportRecord
                            {
                                ExpNumber            = GetStr(row, Col("exp_number"), ""),
                                ExportValueUsd       = expVal,
                                RepatriatedAmountUsd = repat,
                                FreightCharge        = freight,
                                CommissionInsurance  = comm,
                                DestinationCountry   = GetStr(row, Col("destination_country"), "Italy"),
                                ShipmentDate         = shipD,
                                RepatriationDate     = repatD,
                                ClaimedAmountBdt     = rowInc,
                                CertifiedAmountBdt   = rowInc
                            });
                        }

                        double netFob      = totalRepat - (totalFreight + totalComm);
                        double totalIncBdt = netFob * (incPct / 100.0) * exRate;
                        var    fyList      = ExtractFinancialYears(shipDates);

                        _contractsData.Add(new ContractInfo
                        {
                            SheetName           = sheetName,
                            CompanyName         = companyName,
                            IncentiveType       = incType,
                            CircularNo          = circNo,
                            AuditorReference    = audRef,
                            BankName            = bankName,
                            IncentivePercentage = incPct,
                            ExchangeRate        = exRate,
                            TotalExportValueUsd = totalExport,
                            TotalRepatriatedUsd = totalRepat,
                            TotalFreight        = totalFreight,
                            TotalCommission     = totalComm,
                            NetFobUsd           = netFob,
                            TotalIncentiveBdt   = totalIncBdt,
                            ExportCount         = exports.Count,
                            Exports             = exports,
                            FinancialYearsList  = fyList,
                            ApplicationDate     = appDate,
                            ContractNumber      = contractNo,
                            ContractDate        = contractDt,
                            ContractValue       = contractVal
                        });

                        status = $"{exports.Count} EXP(s)  |  BDT {totalIncBdt:N2}";
                    }
                    catch (Exception ex)
                    {
                        status = $"Error: {ex.Message}";
                    }

                    var info = _contractsData.LastOrDefault();
                    _contractRows.Add(new ContractRow
                    {
                        SheetName      = sheetName,
                        CompanyName    = info?.CompanyName ?? companyName,
                        ContractNumber = info?.ContractNumber ?? "-",
                        ContractDate   = info?.ContractDate ?? "-",
                        ContractValue  = info != null ? $"{info.ContractValue:N2}" : "-",
                        Status         = status
                    });
                }

                GenerateBtn.IsEnabled = _contractsData.Count > 0;
                StatusLabel.Text = $"✅ {_contractsData.Count} contract(s) loaded from {wb.Worksheets.Count()} sheet(s).";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load Excel file:\n{ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void GeneratePDFs_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (_contractsData.Count == 0)
            {
                MessageBox.Show("No contract data loaded. Please upload an Excel file first.",
                                "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select Output Directory"
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            string outputDir = dlg.SelectedPath;
            GenerateBtn.IsEnabled = false;
            ProgressBar.Visibility = System.Windows.Visibility.Visible;
            ProgressBar.Value = 0;
            StatusLabel.Text = "Generating documents...";

            try
            {
                var generator = new PdfGeneratorService(
                    CompanyAddressInput.Text,
                    BankBranchInput.Text,
                    AuditFeeNoteInput.Text
                );

                var progress = new Progress<int>(v => ProgressBar.Value = v);

                string zipPath = await Task.Run(() =>
                    generator.GenerateAll(_contractsData, outputDir, progress));

                ProgressBar.Value = 100;
                StatusLabel.Text = $"✅ Done! ZIP saved to: {zipPath}";

                var result = MessageBox.Show(
                    $"PDF generation complete!\n\n" +
                    $"• {_contractsData.Count} Certificate(s)\n" +
                    $"• {_contractsData.Count} Annexure(s)\n" +
                    $"• 1 Audit Details Report\n\n" +
                    $"Saved to:\n{zipPath}\n\nOpen output folder?",
                    "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start("explorer.exe",
                        Path.GetDirectoryName(zipPath)!);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate PDFs:\n{ex.Message}", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                StatusLabel.Text = "Error occurred.";
            }
            finally
            {
                GenerateBtn.IsEnabled = true;
                ProgressBar.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void DownloadTemplate_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "Incentive_Excel_Template.xlsx",
                Title = "Save Template As"
            };
            if (dlg.ShowDialog() == true)
            {
                TemplateService.CreateTemplate(dlg.FileName);
                MessageBox.Show($"Template saved to:\n{dlg.FileName}", "Success",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────────

        private static string GetStr(IXLRow row, int col, string fallback)
        {
            if (col < 1) return fallback;
            var v = row.Cell(col).GetString().Trim();
            return string.IsNullOrEmpty(v) ? fallback : v;
        }

        private static double GetDouble(IXLRow row, int col, double fallback)
        {
            if (col < 1) return fallback;
            try { return row.Cell(col).GetDouble(); }
            catch { return fallback; }
        }

        /// <summary>
        /// Read a cell that might be stored as DateTime, string, or number (Excel serial).
        /// Returns dd/MM/yyyy string, or empty if blank.
        /// </summary>
        private static string FormatDateCell(IXLWorksheet ws, int rowNum, int colNum)
        {
            if (colNum < 1) return "";
            var cell = ws.Cell(rowNum, colNum);
            if (cell.IsEmpty()) return "";

            // Excel stored actual DateTime
            if (cell.DataType == XLDataType.DateTime)
            {
                try { return cell.GetDateTime().ToString("dd/MM/yyyy"); }
                catch { }
            }

            // Try numeric (Excel serial date)
            try
            {
                double serial = cell.GetDouble();
                if (serial > 1000) // sanity: real Excel dates are 40000+
                {
                    var dt = DateTime.FromOADate(serial);
                    return dt.ToString("dd/MM/yyyy");
                }
            }
            catch { }

            // Try raw string parse
            string raw = cell.GetString().Trim();
            if (string.IsNullOrEmpty(raw)) return "";

            string[] fmts = {
                "dd/MM/yyyy", "yyyy-MM-dd", "dd-MM-yyyy",
                "MM/dd/yyyy", "dd.MM.yyyy", "d/M/yyyy",
                "dd/MM/yyyy HH:mm:ss", "yyyy-MM-dd HH:mm:ss",
                "M/d/yyyy", "d/M/yyyy H:mm"
            };
            foreach (var fmt in fmts)
                if (DateTime.TryParseExact(raw, fmt, null,
                    System.Globalization.DateTimeStyles.None, out var dt))
                    return dt.ToString("dd/MM/yyyy");

            if (DateTime.TryParse(raw, out var dt3))
                return dt3.ToString("dd/MM/yyyy");

            return raw;
        }

        private static List<string> ExtractFinancialYears(List<string> dates)
        {
            var fySet = new HashSet<string>();
            foreach (var d in dates)
            {
                DateTime dt;
                bool parsed = DateTime.TryParseExact(d, "dd/MM/yyyy", null,
                    System.Globalization.DateTimeStyles.None, out dt)
                    || DateTime.TryParse(d, out dt);
                if (!parsed) continue;
                int fyStart = dt.Month >= 7 ? dt.Year : dt.Year - 1;
                fySet.Add($"FY-{fyStart}-{fyStart + 1}");
            }
            return fySet.OrderBy(x => x).ToList();
        }
    }
}
