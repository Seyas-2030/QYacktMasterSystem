using System.Windows;
using System.Windows.Controls;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class AdminDashboardWindow : Window
    {
        private readonly AdminDashboardViewModel _vm;

        public AdminDashboardWindow()
        {
            InitializeComponent();
            _vm = new AdminDashboardViewModel();
            DataContext = _vm;
            _vm.LoadAllData();
        }

        // ── Sidebar navigation: toggle panel visibility ──────────────
        private void OnNavClick(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn) return;
            string tag = btn.Tag?.ToString() ?? string.Empty;

            PanelSales   .Visibility = tag == "Sales"    ? Visibility.Visible : Visibility.Collapsed;
            PanelProfit  .Visibility = tag == "Profit"   ? Visibility.Visible : Visibility.Collapsed;
            PanelDebts   .Visibility = tag == "Debts"    ? Visibility.Visible : Visibility.Collapsed;
            PanelTopParts.Visibility = tag == "TopParts" ? Visibility.Visible : Visibility.Collapsed;
            PanelLowStock.Visibility = tag == "LowStock" ? Visibility.Visible : Visibility.Collapsed;
        }
    }
}
