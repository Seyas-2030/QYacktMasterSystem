using DocumentFormat.OpenXml.Bibliography;
using QYachtMaster.Database;
using QYachtMaster.ViewModels;
using QYachtMaster.Views;
using Microsoft.Data.Sqlite;
using System;
using System.Windows;

// تم تعديل مساحة الاسم هنا لتطابق الـ XAML تماماً وتُحل مشكلة InitializeComponent والدوال الثلاث
namespace QYachtMaster.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }


        private void OnOpenSyncClick(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "اختر ملف المخزن / Select Inventory Excel File"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var (existing, newParts) = Inventory.PartsSyncEngine.ParseExcelFile(dlg.FileName);
                var syncWin = new ExcelSyncWindow(existing, newParts);
                syncWin.Owner = this;
                syncWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل قراءة الملف / Failed to parse file: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnOpenAdminClick(object sender, RoutedEventArgs e)
        {
            string username = "admin";
            if (DataContext is MainViewModel vm)
            {
                // حل التحذير الأول CS8600 باستخدام معامل دمج القيم الفارغة ??
                username = vm.CurrentUsername ?? "admin";
            }

            // Look up role in Database
            string role = string.Empty;
            try
            {
                string query = "SELECT Role FROM Users WHERE Username = @user AND IsActive = 1;";
                var param = new SqliteParameter("@user", username);

                // تعريف الكائن كـ قابِل للفراغ (object?) لحل التحذير الثاني
                object? result = DatabaseHelper.ExecuteScalar(query, param);
                if (result != null)
                {
                    role = result.ToString() ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطأ في قاعدة البيانات / Database error: {ex.Message}", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (role != "Admin")
            {
                MessageBox.Show("هذا المستخدم لا يملك صلاحيات الإدارة / This user does not have administrator privileges.", "تم رفض الدخول / Access Denied", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Direct admin access - Passcode requirement removed
            var adminWin = new AdminDashboardWindow();
            adminWin.Owner = this;
            adminWin.ShowDialog();
        }
        private void OnReturnPartClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;

            // Guard: we should only reach here if IsCustomerFound == true
            if (!vm.IsCustomerFound)
            {
                MessageBox.Show(
                    "يرجى البحث عن العميل أولاً / Please search for a customer first.",
                    "No Customer / لا يوجد عميل",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Expose the current customer ID via a public property on the VM
            int    custID   = vm.CurrentCustomerID;
            string custName = vm.CustomerName;
            string user     = vm.CurrentUsername ?? "admin";

            var returnWin = new PartReturnWindow(custID, custName, user)
            {
                Owner = this
            };
            returnWin.ShowDialog();
        }

        private void OnViewDebtStatementClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel vm) return;
            if (!vm.IsCustomerFound) return;

            var ledgerWin = new CustomerDebtLedgerWindow(vm.CurrentCustomerID, vm.CustomerName, vm.SearchPhone, vm.TotalDebt)
            {
                Owner = this
            };
            ledgerWin.ShowDialog();
        }

        private void OnOpenDispatcherClick(object sender, RoutedEventArgs e)
        {
            var boardWindow = new Window
            {
                Title = "📡 Monitor Board / لوحة المراقبة",
                Owner = this,
                Content = new DispatcherBoardView(),
                Width = 1200,
                Height = 760,
                MinWidth = 900,
                MinHeight = 600,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };
            boardWindow.Show();
        }
    }
}
