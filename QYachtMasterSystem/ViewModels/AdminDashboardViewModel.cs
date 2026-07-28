using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using QYachtMaster.MVVM;
using QYachtMaster.Reports;
using QYachtMaster.Views;

namespace QYachtMaster.ViewModels
{
    public class AdminDashboardViewModel : ViewModelBase
    {
        private bool _isAuthenticated = false;
        public bool IsAuthenticated
        {
            get => _isAuthenticated;
            set => SetProperty(ref _isAuthenticated, value);
        }
        private string _selectedReportMode = "monthly";
        private string _selectedTargetPeriod = string.Empty;
        private DateTime? _startDate = DateTime.Today.AddMonths(-1);
        private DateTime? _endDate = DateTime.Today;

        public string SelectedReportMode
        {
            get => _selectedReportMode;
            set
            {
                string resolved = value ?? "monthly";
                if (resolved.Contains("System.Windows.Controls.ComboBoxItem"))
                {
                    var knownModes = new[] { "monthly", "custom" };
                    resolved = knownModes.FirstOrDefault(m =>
                        resolved.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0) ?? "monthly";
                }
                resolved = resolved.Trim().ToLower();

                if (SetProperty(ref _selectedReportMode, resolved))
                {
                    OnPropertyChanged(nameof(ShowPeriodSelector));
                    OnPropertyChanged(nameof(ShowCustomDatePickers));
                    RefreshAvailablePeriods();
                }
            }
        }

        public string SelectedTargetPeriod
        {
            get => _selectedTargetPeriod;
            set
            {
                if (SetProperty(ref _selectedTargetPeriod, value))
                {
                    RefreshSalesInvoices();
                }
            }
        }

        public DateTime? StartDate
        {
            get => _startDate;
            set
            {
                if (SetProperty(ref _startDate, value))
                {
                    RefreshSalesInvoices();
                }
            }
        }

        public DateTime? EndDate
        {
            get => _endDate;
            set
            {
                if (SetProperty(ref _endDate, value))
                {
                    RefreshSalesInvoices();
                }
            }
        }

        public bool ShowPeriodSelector => SelectedReportMode != "custom";
        public bool ShowCustomDatePickers => SelectedReportMode == "custom";

        public ObservableCollection<string> AvailablePeriods { get; } = new();

        // ─── KPI Summary Cards ──────────────────────────────────────────
        private double _totalRevenue;
        private int    _totalInvoices;
        private double _totalDebtOutstanding;
        private int    _lowStockCount;
        private double _totalNetProfit;

        public double TotalRevenue
        {
            get => _totalRevenue;
            set => SetProperty(ref _totalRevenue, value);
        }
        public int TotalInvoices
        {
            get => _totalInvoices;
            set => SetProperty(ref _totalInvoices, value);
        }
        public double TotalDebtOutstanding
        {
            get => _totalDebtOutstanding;
            set => SetProperty(ref _totalDebtOutstanding, value);
        }
        public int LowStockCount
        {
            get => _lowStockCount;
            set => SetProperty(ref _lowStockCount, value);
        }
        public double TotalNetProfit
        {
            get => _totalNetProfit;
            set => SetProperty(ref _totalNetProfit, value);
        }

        // ─── Data Collections ───────────────────────────────────────────
        public ObservableCollection<ReportExporter.DetailedReportRow> SalesInvoices { get; } = new();
        public ObservableCollection<ReportExporter.PartsProfitRow>   ProfitRows    { get; } = new();
        public ObservableCollection<ReportExporter.CustomerDebtRow>  DebtRows      { get; } = new();
        public ObservableCollection<ReportExporter.TopSellingPartRow> TopSellingRows{ get; } = new();
        public ObservableCollection<ReportExporter.LowStockPartRow>  LowStockRows  { get; } = new();

        // ─── Commands ───────────────────────────────────────────────────
        public ICommand ExportSalesCommand        { get; }
        public ICommand ExportProfitCommand       { get; }
        public ICommand ExportDebtsCommand        { get; }
        public ICommand ExportTopSellingCommand   { get; }
        public ICommand ExportLowStockCommand     { get; }
        public ICommand OpenExcelSyncCommand      { get; }

        public AdminDashboardViewModel()
        {
            ExportSalesCommand      = new RelayCommand(_ => ExportSales());
            ExportProfitCommand     = new RelayCommand(_ => ExportAndOpen(ReportExporter.ExportPartsProfitability));
            ExportDebtsCommand      = new RelayCommand(_ => ExportAndOpen(ReportExporter.ExportCustomerDebts));
            ExportTopSellingCommand = new RelayCommand(_ => ExportAndOpen(ReportExporter.ExportTopSellingParts));
            ExportLowStockCommand   = new RelayCommand(_ => ExportAndOpen(ReportExporter.ExportLowStockParts));
            OpenExcelSyncCommand    = new RelayCommand(_ => OpenExcelSync());
        }

        private void ExportSales()
        {
            try
            {
                if (SelectedReportMode == "custom")
                {
                    if (!StartDate.HasValue || !EndDate.HasValue)
                    {
                        MessageBox.Show("يرجى تحديد تاريخ البداية والنهاية أولاً!\nPlease select both start and end dates first.",
                                        "تنبيه / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    string startStr = StartDate.Value.ToString("yyyy-MM-dd");
                    string endStr = EndDate.Value.ToString("yyyy-MM-dd");
                    ExportAndOpen(() => ReportExporter.ExportCustomDetailedReport(startStr, endStr));
                }
                else
                {
                    // Automatic sequencing for monthly report
                    int year = 0, month = 0;
                    ExportAndOpen(() => ReportExporter.ExportMonthlyReportAuto(out year, out month));
                    
                    // Refresh available periods on UI to make the new export immediately listed
                    RefreshAvailablePeriods();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطأ أثناء تصدير التقرير / Error exporting report: {ex.Message}",
                                "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void LoadAllData()
        {
            RefreshAvailablePeriods();
            RefreshProfitData();
            RefreshDebtData();
            RefreshTopSelling();
            RefreshLowStock();
            RefreshKpis();
        }

        private void RefreshAvailablePeriods()
        {
            AvailablePeriods.Clear();
            if (SelectedReportMode != "custom")
            {
                var periods = ReportExporter.FetchAvailablePeriods(SelectedReportMode);
                foreach (var p in periods)
                {
                    AvailablePeriods.Add(p);
                }
                if (AvailablePeriods.Count > 0)
                {
                    SelectedTargetPeriod = AvailablePeriods[0];
                }
                else
                {
                    SelectedTargetPeriod = string.Empty;
                    RefreshSalesInvoices();
                }
            }
            else
            {
                RefreshSalesInvoices();
            }
        }

        private void RefreshSalesInvoices()
        {
            SalesInvoices.Clear();

            List<ReportExporter.DetailedReportRow> data = new();
            if (SelectedReportMode == "custom")
            {
                if (StartDate.HasValue && EndDate.HasValue)
                {
                    string startStr = StartDate.Value.ToString("yyyy-MM-dd");
                    string endStr = EndDate.Value.ToString("yyyy-MM-dd");
                    data = ReportExporter.FetchCustomDetailedReport(startStr, endStr);
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(SelectedTargetPeriod))
                {
                    data = ReportExporter.FetchDetailedReport(SelectedReportMode, SelectedTargetPeriod);
                }
            }

            foreach (var r in data)
            {
                SalesInvoices.Add(r);
            }

            // Update KPIs dynamically based on current invoices
            double rev = 0, profit = 0;
            foreach (var r in SalesInvoices)
            {
                rev += r.GrandTotal;
                profit += r.NetProfit;
            }
            TotalRevenue = rev;
            TotalNetProfit = profit;
            TotalInvoices = SalesInvoices.Count;
        }

        private void RefreshProfitData()
        {
            ProfitRows.Clear();
            foreach (var r in ReportExporter.FetchPartsProfitability())
                ProfitRows.Add(r);
        }

        private void RefreshDebtData()
        {
            DebtRows.Clear();
            foreach (var r in ReportExporter.FetchCustomerDebts())
                DebtRows.Add(r);
        }

        private void RefreshTopSelling()
        {
            TopSellingRows.Clear();
            foreach (var r in ReportExporter.FetchTopSellingParts())
                TopSellingRows.Add(r);
        }

        private void RefreshLowStock()
        {
            LowStockRows.Clear();
            foreach (var r in ReportExporter.FetchLowStockParts())
                LowStockRows.Add(r);
        }

        private void RefreshKpis()
        {
            // Debt comes from DebtRows (customer debt table)
            double debt = 0;
            foreach (var r in DebtRows)
                debt += r.TotalDebt;

            TotalDebtOutstanding = debt;
            LowStockCount        = LowStockRows.Count;
            // Revenue/Profit/InvoiceCount are updated live in RefreshSalesInvoices
        }

        private void ExportAndOpen(Func<string> exportFunc)
        {
            try
            {
                string path = exportFunc();
                MessageBox.Show($"تم تصدير التقرير بنجاح!\nExport saved:\n{path}",
                    "تصدير ناجح / Export Success", MessageBoxButton.OK, MessageBoxImage.Information);
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل التصدير / Export failed: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenExcelSync()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title  = "اختر ملف المخزن / Select Inventory Excel File"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var (existing, newParts) = Inventory.PartsSyncEngine.ParseExcelFile(dlg.FileName);
                var syncWin = new ExcelSyncWindow(existing, newParts);
                syncWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل قراءة الملف / Failed to parse file: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
