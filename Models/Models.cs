using System.Collections.Generic;
using System.ComponentModel;

namespace ExportIncentiveApp.Models
{
    public class ContractInfo
    {
        public string SheetName { get; set; } = "";
        public string CompanyName { get; set; } = "";
        public string IncentiveType { get; set; } = "";
        public string CircularNo { get; set; } = "";
        public string AuditorReference { get; set; } = "";
        public string BankName { get; set; } = "";
        public double IncentivePercentage { get; set; }
        public double ExchangeRate { get; set; }
        public double TotalExportValueUsd { get; set; }
        public double TotalRepatriatedUsd { get; set; }
        public double TotalFreight { get; set; }
        public double TotalCommission { get; set; }
        public double NetFobUsd { get; set; }
        public double TotalIncentiveBdt { get; set; }
        public int ExportCount { get; set; }
        public List<ExportRecord> Exports { get; set; } = new();
        public List<string> FinancialYearsList { get; set; } = new();
        public string ApplicationDate { get; set; } = "";
        public string ContractNumber { get; set; } = "";
        public string ContractDate { get; set; } = "";
        public double ContractValue { get; set; }
    }

    public class ExportRecord
    {
        public string ExpNumber { get; set; } = "";
        public double ExportValueUsd { get; set; }
        public double RepatriatedAmountUsd { get; set; }
        public double FreightCharge { get; set; }
        public double CommissionInsurance { get; set; }
        public string DestinationCountry { get; set; } = "";
        public string ShipmentDate { get; set; } = "";
        public string RepatriationDate { get; set; } = "";
        public double ClaimedAmountBdt { get; set; }
        public double CertifiedAmountBdt { get; set; }
    }

    // DataGrid binding model (editable)
    public class ContractRow : INotifyPropertyChanged
    {
        private string _contractNumber = "";
        private string _contractDate = "";
        private string _contractValue = "";

        public string SheetName { get; set; } = "";
        public string CompanyName { get; set; } = "";
        public string Status { get; set; } = "";

        public string ContractNumber
        {
            get => _contractNumber;
            set { _contractNumber = value; OnPropertyChanged(nameof(ContractNumber)); }
        }
        public string ContractDate
        {
            get => _contractDate;
            set { _contractDate = value; OnPropertyChanged(nameof(ContractDate)); }
        }
        public string ContractValue
        {
            get => _contractValue;
            set { _contractValue = value; OnPropertyChanged(nameof(ContractValue)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
