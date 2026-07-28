using System;
using System.Collections.Generic;
using System.Windows;
using QYachtMaster.Reports;

namespace QYachtMaster.Views
{
    public partial class ReportFilterWindow : Window
    {
        public string SelectedPeriod { get; private set; } = string.Empty;
        public bool IsCustomRange { get; private set; } = false;
        public string CustomFromDate { get; private set; } = string.Empty;
        public string CustomToDate { get; private set; } = string.Empty;

        public ReportFilterWindow(string periodType)
        {
            InitializeComponent();

            // Populate combo box with monthly periods found in DB
            List<string> periods = ReportExporter.FetchAvailablePeriods(periodType);
            foreach (var p in periods)
            {
                PeriodCombo.Items.Add(p);
            }

            if (PeriodCombo.Items.Count > 0)
            {
                PeriodCombo.SelectedIndex = 0;
            }

            // Default custom range: current month
            DpFrom.SelectedDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DpTo.SelectedDate   = DateTime.Now;
        }

        private void OnModeChanged(object sender, RoutedEventArgs e)
        {
            if (PanelMonthly == null || PanelCustom == null) return;

            bool isMonthly = RbMonthly.IsChecked == true;
            PanelMonthly.Visibility = isMonthly ? Visibility.Visible : Visibility.Collapsed;
            PanelCustom.Visibility  = isMonthly ? Visibility.Collapsed : Visibility.Visible;
        }

        private void OnExportClick(object sender, RoutedEventArgs e)
        {
            if (RbMonthly.IsChecked == true)
            {
                // Monthly mode
                if (PeriodCombo.SelectedItem == null)
                {
                    MessageBox.Show("الرجاء اختيار شهر / Please select a month.",
                                    "تنبيه / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                SelectedPeriod = PeriodCombo.SelectedItem.ToString() ?? string.Empty;
                IsCustomRange  = false;
            }
            else
            {
                // Custom date range mode
                if (DpFrom.SelectedDate == null || DpTo.SelectedDate == null)
                {
                    MessageBox.Show("الرجاء تحديد تاريخ البدء والانتهاء / Please select both start and end dates.",
                                    "تنبيه / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (DpTo.SelectedDate < DpFrom.SelectedDate)
                {
                    MessageBox.Show("تاريخ الانتهاء يجب أن يكون بعد تاريخ البدء / End date must be after start date.",
                                    "تنبيه / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                CustomFromDate = DpFrom.SelectedDate.Value.ToString("yyyy-MM-dd");
                CustomToDate   = DpTo.SelectedDate.Value.ToString("yyyy-MM-dd");
                IsCustomRange  = true;
            }

            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
