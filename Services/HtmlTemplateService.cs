using System;
using System.Collections.Generic;
using System.Text;
using ExportIncentiveApp.Models;

namespace ExportIncentiveApp.Services
{
    /// <summary>Builds HTML strings for Certificate, Annexure, and Bill</summary>
    public static class HtmlTemplateService
    {
        private static string FC(double v) => $"{v:N2}";

        // ── Certificate ───────────────────────────────────────────────────────────
        public static string BuildCertificate(ContractInfo c, string bankRef, string serialNo)
        {
            string amountWords = NumberToWords(c.TotalIncentiveBdt);
            string contractDateFmt = FormatDateDisplay(c.ContractDate);

            return $@"<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<style>
  body {{ font-family: 'Times New Roman', serif; font-size: 13px; line-height: 1.5; color:#000; margin:40px; }}
  .ref {{ font-size:12px; line-height:1.2; }}
  .title {{ text-align:center; margin:20px 0; }}
  .title h1 {{ font-size:20px; text-decoration:underline; margin:0; text-transform:uppercase; }}
  .title h2 {{ font-size:16px; margin:5px 0 0; }}
  .title h3 {{ font-size:14px; font-weight:normal; margin:2px 0; }}
  .body {{ text-align:justify; margin-top:30px; }}
  .footer {{ margin-top:40px; }}
  .sig {{ float:right; width:260px; margin-top:80px; }}
  .addr {{ font-size:11px; line-height:1.3; }}
</style></head><body>
  <table width='100%'><tr><td class='ref'>
    Auditor's Ref: {c.AuditorReference}<br>
    Bank Ref: {bankRef}<br>
    Serial No: {serialNo}
  </td></tr></table>

  <div class='title'>
    <h1>Certification</h1>
    <h2>{c.IncentiveType}</h2>
    <h3>As per {c.CircularNo}</h3>
  </div>

  <div class='body'>
    <p>
      Applicant <strong>{c.CompanyName}</strong> exported RMG through
      Export LC/Contract number <strong>({c.ContractNumber})</strong>
      dated <strong>({contractDateFmt})</strong> for
      USD <strong>({FC(c.ContractValue)})</strong>
      vide <strong>{c.ExportCount}</strong> Nos. EXP numbers of {c.BankName}
      containing Export Value of <strong>USD {FC(c.TotalExportValueUsd)}</strong>
      against which <strong>USD {FC(c.TotalRepatriatedUsd)}</strong>
      has been Repatriated and {c.IncentiveType.ToLower()} has been claimed for
      <strong>BDT {FC(c.TotalIncentiveBdt)}</strong>.
    </p>
    <p>
      After audit, the amount payable as {c.IncentiveType.ToLower()}
      @ {c.IncentivePercentage}% works out at
      <strong>BDT {FC(c.TotalIncentiveBdt)}</strong>
      (Taka {amountWords} Only),
      which is hereby certified to be true and fair.
    </p>
  </div>

  <div class='footer'>
    <strong>FY-2024-2025: BDT {FC(c.TotalIncentiveBdt)}</strong>
    <div class='sig'>
      <strong>Signature: _________________________</strong><br><br>
      <div style='margin-left:50px;'>
        <strong>Md. Mahamud Hosain FCA</strong><br>
        Managing Partner<br>Mahamud Sabuj &amp; Co.<br>Chartered Accountants<br>
        <div class='addr'>Fare Diya Complex<br>House: 11/8/E(7th Floor)<br>Free School Street<br>Panthapath, Dhaka-1205.</div>
      </div>
    </div>
    <div style='margin-top:140px;'>Dhaka, {DateTime.Now:dd/MM/yyyy}<br>Enclosure: Annexure-A.</div>
  </div>
</body></html>";
        }

        // ── Annexure ──────────────────────────────────────────────────────────────
        public static string BuildAnnexure(ContractInfo c, string bankRef, string serialNo)
        {
            string contractDateFmt = FormatDateDisplay(c.ContractDate);
            var fyRows = new StringBuilder();
            foreach (var fy in c.FinancialYearsList) fyRows.AppendLine($"<div>{fy}</div>");

            var exportRows = new StringBuilder();
            foreach (var exp in c.Exports)
            {
                exportRows.AppendLine($@"<tr>
  <td>{exp.ExpNumber}</td>
  <td class='r'>{FC(exp.ExportValueUsd)}</td>
  <td>{exp.DestinationCountry}</td>
  <td>{exp.ShipmentDate}</td>
  <td>{exp.RepatriationDate}</td>
  <td class='r'>{FC(exp.RepatriatedAmountUsd)}</td>
  <td class='r'>{FC(exp.ClaimedAmountBdt)}</td>
  <td class='r'>{FC(exp.CertifiedAmountBdt)}</td>
</tr>");
            }

            string commCell = c.TotalCommission > 0 ? FC(c.TotalCommission) : "-";

            return $@"<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<style>
  body {{ font-family:'Times New Roman',serif; font-size:11px; line-height:1.3; color:#000; margin:20px; }}
  table {{ width:100%; border-collapse:collapse; margin-bottom:15px; }}
  th,td {{ border:1px solid #000; padding:4px; text-align:center; }}
  th {{ background:#f2f2f2; font-weight:bold; }}
  .r {{ text-align:right; }} .b {{ font-weight:bold; }} .l {{ text-align:left; }}
  .sh {{ font-weight:bold; margin:10px 0 4px; }}
</style></head><body>
  <table style='border:none' display='flex'><tr>
    <td style='border:none; width:50%; text-align:left;'>
      Auditor's Ref: {c.AuditorReference}<br>Bank Ref: {bankRef}<br>Serial No: {serialNo}
    </td>
    <td style='border:none; text-align:right; width:50%;'>
      <strong style='font-size:14px;'>Annexure-A</strong><br><br>
      Application Date: {c.ApplicationDate}<br>
      {fyRows}
    </td>
  </tr></table>

  <div class='b' style='text-transform:uppercase;'>{c.CompanyName}</div>
  <div>{c.IncentiveType}<br>As per {c.CircularNo}</div>

  <div class='sh'>LC/Contract Information:</div>
  <table style='width:60%'>
    <tr><th>LC/Contract number</th><th>Date</th><th>Amount (USD)</th></tr>
    <tr><td>{c.ContractNumber}</td><td>{contractDateFmt}</td><td class='r'>{FC(c.ContractValue)}</td></tr>
    <tr class='b'><td colspan='2'>Total</td><td class='r'>{FC(c.ContractValue)}</td></tr>
  </table>

  <div class='sh'>Export Information:</div>
  <table>
    <tr>
      <th>EXP number</th><th>Export Value (USD)</th><th>Country</th>
      <th>Shipment Date</th><th>Repatriation Date</th>
      <th>Repatriated (USD)</th><th>Claimed (BDT)</th><th>Certified (BDT)</th>
    </tr>
    {exportRows}
    <tr class='b'>
      <td class='l'>Total</td>
      <td class='r'>{FC(c.TotalExportValueUsd)}</td>
      <td colspan='3'></td>
      <td class='r'>{FC(c.TotalRepatriatedUsd)}</td>
      <td class='r'>{FC(c.TotalIncentiveBdt)}</td>
      <td class='r'>{FC(c.TotalIncentiveBdt)}</td>
    </tr>
  </table>

  <div class='sh'>Calculation of Cash Subsidy:</div>
  <table>
    <tr>
      <th>Total Repatriated (USD)</th><th>Freight (C&amp;F)</th><th>Commission/Insurance</th>
      <th>Net FOB</th><th>Amount</th><th>Rate</th><th>Conv. Rate</th><th>Total</th>
    </tr>
    <tr>
      <td rowspan='2'>{FC(c.TotalRepatriatedUsd)}</td>
      <td rowspan='2'>{FC(c.TotalFreight)}</td>
      <td rowspan='2'>{commCell}</td>
      <td rowspan='2'>{FC(c.NetFobUsd)}</td>
      <td class='r'>{FC(c.NetFobUsd)}</td>
      <td>{c.IncentivePercentage}%</td>
      <td>{c.ExchangeRate:F2}</td>
      <td class='r b'>{FC(c.TotalIncentiveBdt)}</td>
    </tr>
    <tr>
      <td colspan='3' class='r b'>BDT</td>
      <td class='r b'>{FC(c.TotalIncentiveBdt)}</td>
    </tr>
  </table>
  <div style='font-size:10px; font-style:italic;'>(Note: Applicable rate as per FE Circular No. 12 Dated 30/06/2024)</div>
</body></html>";
        }

        // ── Bill ─────────────────────────────────────────────────────────────────
        public static string BuildBill(
            List<BillContractRow> rows,
            double totalApplied, double totalPayable, double totalAuditFee,
            string incentiveTypePlural, string auditFeeNote)
        {
            var now = DateTime.Now;
            var tbody = new StringBuilder();
            int idx = 1;
            foreach (var r in rows)
            {
                tbody.AppendLine($@"<tr>
  <td>{idx++}</td>
  <td>{r.BankName}, {r.BankBranch}</td>
  <td>{r.CompanyName}<br>{r.CompanyAddress}</td>
  <td>{r.AuditorReference}<br>{r.BankReference}</td>
  <td>{r.ContractNumber}</td>
  <td>{r.ContractDate}</td>
  <td>USD</td>
  <td class='r'>{FC(r.ExportValueUsd)}</td>
  <td class='r'>{FC(r.RepatriatedUsd)}</td>
  <td class='r'>{FC(r.AppliedAmount)}</td>
  <td class='r'>{FC(r.PayableAmount)}</td>
  <td class='r'>{FC(r.AuditFee)}</td>
</tr>");
            }

            double vat = totalAuditFee * 0.15;
            double totalWithVat = totalAuditFee * 1.15;

            return $@"<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<style>
  body {{ font-family:'Segoe UI',sans-serif; font-size:12px; margin:40px; color:#000; }}
  .ref {{ text-align:right; font-weight:bold; }}
  .title {{ text-align:center; font-size:16px; font-weight:bold; text-decoration:underline; margin-bottom:20px; }}
  table {{ width:100%; border-collapse:collapse; }}
  th,td {{ border:1px solid #000; padding:6px; text-align:center; vertical-align:middle; }}
  th {{ background:#f2f2f2; font-weight:bold; }}
  .r {{ text-align:right; }} .b {{ font-weight:bold; }} .nb {{ border:none!important; }}
  .sig {{ float:right; width:260px; margin-top:80px; }}
</style></head><body>
  <div class='ref'>Ref: MSC/UBPLC/CIAudit/{now:yyyy}/{now:MM}/{now:dd}/001</div>
  <div class='title'>Details of Audited Cases</div>

  <table>
    <tr>
      <th rowspan='2'>Sl.</th>
      <th rowspan='2'>Bank &amp; Branch</th>
      <th rowspan='2'>Exporter Name &amp; Address</th>
      <th rowspan='2'>Ref No.</th>
      <th rowspan='2'>LC/Contract</th>
      <th rowspan='2'>Date</th>
      <th rowspan='2'>Currency</th>
      <th rowspan='2'>Export Value</th>
      <th rowspan='2'>Repatriated</th>
      <th colspan='2'>{incentiveTypePlural} (Taka)</th>
      <th rowspan='2'>Audit Fee (Taka)</th>
    </tr>
    <tr><th>Applied</th><th>Payable</th></tr>
    {tbody}
    <tr class='b'>
      <td colspan='9' style='text-align:center;'>Total:</td>
      <td class='r'>{FC(totalApplied)}</td>
      <td class='r'>{FC(totalPayable)}</td>
      <td class='r'>{FC(totalAuditFee)}</td>
    </tr>
    <tr class='b'>
      <td colspan='8' class='nb'></td>
      <td colspan='3' style='text-align:left;'>VAT on audit fee @ 15%:</td>
      <td class='r'>{FC(vat)}</td>
    </tr>
    <tr class='b'>
      <td colspan='8' class='nb'></td>
      <td colspan='3' style='text-align:left;'>Total including VAT:</td>
      <td class='r'>{FC(totalWithVat)}</td>
    </tr>
  </table>

  <p style='font-size:11px; font-style:italic;'>* {auditFeeNote}</p>
  <br/><br/><br/>
  <div class='sig'>
    <strong>Signature: _________________________</strong><br><br>
    <div style='margin-left:50px;'>
      <strong>Md. Mahamud Hosain FCA</strong><br>
      Managing Partner<br>Mahamud Sabuj &amp; Co.<br>Chartered Accountants<br>
      <div style='font-size:11px; line-height:1.3;'>Fare Diya Complex<br>House: 11/8/E(7th Floor)<br>Free School Street<br>Panthapath, Dhaka-1205.</div>
    </div>
  </div>
</body></html>";
        }

        // ── Helpers ───────────────────────────────────────────────────────────────
        private static string FormatDateDisplay(string raw)
        {
            string[] fmts = { "yyyy-MM-dd", "dd/MM/yyyy", "dd-MM-yyyy", "MM/dd/yyyy" };
            foreach (var fmt in fmts)
                if (DateTime.TryParseExact(raw, fmt, null, System.Globalization.DateTimeStyles.None, out var dt))
                    return dt.ToString("dd/MM/yyyy");
            return raw;
        }

        public static string NumberToWords(double num)
        {
            string[] ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };
            string[] teens = { "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
            string[] tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

            string BelowThousand(int n)
            {
                if (n == 0) return "";
                if (n < 10) return ones[n];
                if (n < 20) return teens[n - 10];
                if (n < 100) return tens[n / 10] + (n % 10 != 0 ? " " + ones[n % 10] : "");
                return ones[n / 100] + " Hundred" + (n % 100 != 0 ? " " + BelowThousand(n % 100) : "");
            }

            if (num == 0) return "Zero";
            long intPart = (long)num;
            var result = new StringBuilder();

            if (intPart >= 10_000_000) { result.Append(BelowThousand((int)(intPart / 10_000_000)) + " Crore "); intPart %= 10_000_000; }
            if (intPart >= 100_000) { result.Append(BelowThousand((int)(intPart / 100_000)) + " Lakh "); intPart %= 100_000; }
            if (intPart >= 1_000) { result.Append(BelowThousand((int)(intPart / 1_000)) + " Thousand "); intPart %= 1_000; }
            if (intPart > 0) result.Append(BelowThousand((int)intPart));

            int paisa = (int)Math.Round((num - (long)num) * 100);
            if (paisa > 0) result.Append($" and {paisa}/100");

            return result.ToString().Trim() + " Only";
        }
    }

    public class BillContractRow
    {
        public string CompanyName { get; set; } = "";
        public string CompanyAddress { get; set; } = "";
        public string ContractNumber { get; set; } = "";
        public string ContractDate { get; set; } = "";
        public string AuditorReference { get; set; } = "";
        public string BankReference { get; set; } = "";
        public string BankName { get; set; } = "";
        public string BankBranch { get; set; } = "";
        public double ExportValueUsd { get; set; }
        public double RepatriatedUsd { get; set; }
        public double AppliedAmount { get; set; }
        public double PayableAmount { get; set; }
        public double AuditFee { get; set; }
    }
}
