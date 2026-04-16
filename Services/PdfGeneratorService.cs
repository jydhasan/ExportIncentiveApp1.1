using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using DinkToPdf;
using DinkToPdf.Contracts;
using ExportIncentiveApp.Models;

using DinkOrientation = DinkToPdf.Orientation;

namespace ExportIncentiveApp.Services
{
    public class PdfGeneratorService
    {
        private readonly string _companyAddress;
        private readonly string _bankBranch;
        private readonly string _auditFeeNote;

        // Lazy-init so it's created on the background thread
        private static IConverter? _converter;
        private static readonly object _lock = new();

        public PdfGeneratorService(string companyAddress, string bankBranch, string auditFeeNote)
        {
            _companyAddress = companyAddress;
            _bankBranch = bankBranch;
            _auditFeeNote = auditFeeNote;
        }

        private static IConverter GetConverter()
        {
            if (_converter != null) return _converter;
            lock (_lock)
            {
                if (_converter != null) return _converter;

                // Load the native DLL explicitly from EXE directory
                string exeDir = AppContext.BaseDirectory;
                string dllPath = Path.Combine(exeDir, "libwkhtmltox.dll");

                if (!File.Exists(dllPath))
                    throw new FileNotFoundException(
                        $"libwkhtmltox.dll not found at:\n{dllPath}\n\n" +
                        "Please place libwkhtmltox.dll in the same folder as the EXE.");

                NativeLibrary.Load(dllPath);
                _converter = new SynchronizedConverter(new PdfTools());
                return _converter;
            }
        }

        public string GenerateAll(List<ContractInfo> contracts, string outputDir,
                                  IProgress<int> progress)
        {
            var converter = GetConverter();
            var pdfFiles = new List<string>();
            var billRows = new List<BillContractRow>();
            double totalApplied = 0, totalPayable = 0, totalAuditFee = 0;
            var incentiveTypes = new List<string>();

            for (int i = 0; i < contracts.Count; i++)
            {
                progress?.Report((int)((double)i / contracts.Count * 85));
                var c = contracts[i];

                string bankRef   = $"UBP/LOCAL/ASSISTANCE/{DateTime.Now.Year}/{i + 1:D2}";
                string serialNo  = $"{i + 1:D2}";
                string auditorRef = $"MSC-UBPLC-LO-SES-{DateTime.Now.Year}-{395 + i:D3}";

                // Certificate
                string certHtml  = HtmlTemplateService.BuildCertificate(c, bankRef, serialNo);
                string certPath  = Path.Combine(outputDir, $"Certificate_{Sanitize(c.CompanyName)}_{i + 1}.pdf");
                ConvertToPdf(converter, certHtml, certPath, DinkOrientation.Portrait);
                pdfFiles.Add(certPath);

                // Annexure
                string annexHtml = HtmlTemplateService.BuildAnnexure(c, bankRef, serialNo);
                string annexPath = Path.Combine(outputDir, $"Annexure_{Sanitize(c.CompanyName)}_{i + 1}.pdf");
                ConvertToPdf(converter, annexHtml, annexPath, DinkOrientation.Landscape);
                pdfFiles.Add(annexPath);

                double auditFee  = CalcAuditFee(c.TotalIncentiveBdt);
                double applied   = c.TotalIncentiveBdt * 1.01;
                double payable   = c.TotalIncentiveBdt;
                totalAuditFee   += auditFee;
                totalApplied    += applied;
                totalPayable    += payable;
                incentiveTypes.Add(c.IncentiveType);

                billRows.Add(new BillContractRow
                {
                    CompanyName      = c.CompanyName,
                    CompanyAddress   = _companyAddress,
                    ContractNumber   = c.ContractNumber,
                    ContractDate     = FormatDateDisplay(c.ContractDate),
                    AuditorReference = auditorRef,
                    BankReference    = bankRef,
                    BankName         = c.BankName,
                    BankBranch       = _bankBranch,
                    ExportValueUsd   = c.TotalExportValueUsd,
                    RepatriatedUsd   = c.TotalRepatriatedUsd,
                    AppliedAmount    = applied,
                    PayableAmount    = payable,
                    AuditFee         = auditFee
                });
            }

            progress?.Report(88);

            string incPlural = incentiveTypes.Any()
                ? incentiveTypes.GroupBy(x => x).OrderByDescending(g => g.Count()).First().Key
                : "Incentive";

            string billHtml = HtmlTemplateService.BuildBill(
                billRows, totalApplied, totalPayable, totalAuditFee,
                incPlural, _auditFeeNote);
            string billPath = Path.Combine(outputDir, "Audit_Details_Report.pdf");
            ConvertToPdf(converter, billHtml, billPath, DinkOrientation.Landscape);
            pdfFiles.Add(billPath);

            progress?.Report(95);

            string zipPath = Path.Combine(outputDir,
                $"Export_Documents_{DateTime.Now:yyyyMMdd_HHmmss}.zip");

            using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                foreach (var f in pdfFiles)
                    if (File.Exists(f))
                        zip.CreateEntryFromFile(f, Path.GetFileName(f));

            progress?.Report(100);
            return zipPath;
        }

        private static void ConvertToPdf(IConverter converter, string html,
                                         string outputPath, DinkOrientation orientation)
        {
            var doc = new HtmlToPdfDocument
            {
                GlobalSettings = new GlobalSettings
                {
                    ColorMode  = ColorMode.Color,
                    Orientation = orientation,
                    PaperSize  = PaperKind.A4,
                    Out        = outputPath
                },
                Objects =
                {
                    new ObjectSettings
                    {
                        HtmlContent = html,
                        WebSettings = new WebSettings { DefaultEncoding = "utf-8" }
                    }
                }
            };
            converter.Convert(doc);
        }

        private static double CalcAuditFee(double bdt) =>
            bdt <= 500_000 ? 4_000 : bdt <= 1_000_000 ? 5_000 : 7_000;

        private static string Sanitize(string name) =>
            string.Concat(name.Select(c =>
                Path.GetInvalidFileNameChars().Contains(c) ? '_' : c)).Trim();

        private static string FormatDateDisplay(string raw)
        {
            string[] fmts = { "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy" };
            foreach (var fmt in fmts)
                if (DateTime.TryParseExact(raw, fmt, null,
                    System.Globalization.DateTimeStyles.None, out var dt))
                    return dt.ToString("dd/MM/yyyy");
            return raw;
        }
    }
}
