using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Microsoft.Data.Sqlite;
using QYachtMaster.MVVM;
using QYachtMaster.Database;
using QYachtMaster.Models;
using QYachtMaster.Views;

namespace QYachtMaster.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private string _searchPhone = string.Empty;
        private string _customerName = string.Empty;
        private string _customerAddress = string.Empty;
        private string _workPhone = string.Empty;
        private string _otherPhone = string.Empty;
        private string _statusMessage = string.Empty;
        private string _statusBrush = "Transparent";
        private double _totalDebt = 0.0;
        private bool _isFieldsEnabled = false;
        private bool _isAcceptCashVisible = false;
        private bool _isJobCardButtonEnabled = false;
        private bool _isCustomerFound = false;
        private int _currentCustomerId = 0;

        /// <summary>Public read for code-behind (e.g. PartReturnWindow). Do not bind in XAML.</summary>
        public int CurrentCustomerID => _currentCustomerId;

        // Current User context (seeded or logged in)
        public string CurrentUsername { get; set; } = "admin";

        public string SearchPhone
        {
            get => _searchPhone;
            set
            {
                if (SetProperty(ref _searchPhone, value))
                {
                    OnPhoneChanged();
                }
            }
        }

        public string CustomerName
        {
            get => _customerName;
            set => SetProperty(ref _customerName, value);
        }

        public string CustomerAddress
        {
            get => _customerAddress;
            set => SetProperty(ref _customerAddress, value);
        }

        public string WorkPhone
        {
            get => _workPhone;
            set => SetProperty(ref _workPhone, value);
        }

        public string OtherPhone
        {
            get => _otherPhone;
            set => SetProperty(ref _otherPhone, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public string StatusBrush
        {
            get => _statusBrush;
            set => SetProperty(ref _statusBrush, value);
        }

        public double TotalDebt
        {
            get => _totalDebt;
            set
            {
                if (SetProperty(ref _totalDebt, value))
                {
                    OnPropertyChanged(nameof(HasDebt));
                    OnPropertyChanged(nameof(DebtWarningText));
                }
            }
        }

        public bool HasDebt => TotalDebt > 0;
        public string DebtWarningText => $"⚠️ Outstanding Debt: {TotalDebt} QAR / مديونية: {TotalDebt} ر.ق";

        public bool IsFieldsEnabled
        {
            get => _isFieldsEnabled;
            set => SetProperty(ref _isFieldsEnabled, value);
        }

        public bool IsAcceptCashVisible
        {
            get => _isAcceptCashVisible;
            set => SetProperty(ref _isAcceptCashVisible, value);
        }

        public bool IsJobCardButtonEnabled
        {
            get => _isJobCardButtonEnabled;
            set => SetProperty(ref _isJobCardButtonEnabled, value);
        }

        public bool IsCustomerFound
        {
            get => _isCustomerFound;
            set => SetProperty(ref _isCustomerFound, value);
        }

        public ICommand RegisterOrOpenCommand { get; }
        public ICommand AcceptCashCommand { get; }

        public MainViewModel()
        {
            RegisterOrOpenCommand = new RelayCommand(OnRegisterOrOpen);
            AcceptCashCommand = new RelayCommand(OnAcceptCash);
            ResetState();
        }

        private void ResetState()
        {
            CustomerName = string.Empty;
            CustomerAddress = string.Empty;
            WorkPhone = string.Empty;
            OtherPhone = string.Empty;
            TotalDebt = 0.0;
            StatusMessage = string.Empty;
            StatusBrush = "Transparent";
            IsFieldsEnabled = false;
            IsAcceptCashVisible = false;
            IsJobCardButtonEnabled = false;
            IsCustomerFound = false;
            _currentCustomerId = 0;
        }

        private static string NormalizePhone(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            // Strip all non-digit characters
            string digits = Regex.Replace(input, @"\D", "");
            // Return only the last 8 digits — country code prefixes are ignored
            // e.g. +97412345678 → 12345678, 012345678 → 12345678
            if (digits.Length > 8)
                digits = digits.Substring(digits.Length - 8);
            return digits;
        }

        private void OnPhoneChanged()
        {
            string clean = NormalizePhone(SearchPhone);
            // Trigger search once at least 4 digits entered (after stripping prefix)
            if (clean.Length >= 8)
            {
                QueryCustomer(clean);
            }
            else
            {
                ResetState();
            }
        }

        private void QueryCustomer(string cleanPhone)
        {
            // cleanPhone is already the last 8 digits.
            // For DB comparison: strip all non-digits from stored numbers and compare last 8 digits.
            // SQLite: SUBSTR(col, -8) gives last 8 chars after stripping non-digits via REPLACE chain.
            string query = @"
                SELECT * FROM Customers 
                WHERE SUBSTR(REPLACE(REPLACE(REPLACE(REPLACE(MobilePhone, ' ', ''), '-', ''), '+', ''), '(', ''), -8) = @clean
                   OR SUBSTR(REPLACE(REPLACE(REPLACE(REPLACE(WorkPhone,  ' ', ''), '-', ''), '+', ''), '(', ''), -8) = @clean
                   OR SUBSTR(REPLACE(REPLACE(REPLACE(REPLACE(OtherPhone, ' ', ''), '-', ''), '+', ''), '(', ''), -8) = @clean
                LIMIT 1;";

            var cmdParams = new[]
            {
                new SqliteParameter("@clean", cleanPhone)
            };

            var customer = DatabaseHelper.ExecuteSingleQuery(query, reader => new Customer
            {
                CustomerID = Convert.ToInt32(reader["CustomerID"]),
                Name = reader["Name"].ToString()!,
                MobilePhone = reader["MobilePhone"].ToString()!,
                WorkPhone = reader["WorkPhone"]?.ToString(),
                OtherPhone = reader["OtherPhone"]?.ToString(),
                Address = reader["Address"]?.ToString(),
                TotalDebt = Convert.ToDouble(reader["TotalDebt"]),
                CreatedAt = reader["CreatedAt"].ToString()!
            }, cmdParams);

            if (customer != null)
            {
                // Scenario B - Existing Customer
                _currentCustomerId = customer.CustomerID;
                CustomerName = customer.Name;
                CustomerAddress = customer.Address ?? string.Empty;
                WorkPhone = customer.WorkPhone ?? string.Empty;
                OtherPhone = customer.OtherPhone ?? string.Empty;
                TotalDebt = customer.TotalDebt;
                IsCustomerFound = true;
                IsFieldsEnabled = false;

                StatusMessage = "✔️ Registered Customer / عميل مسجل";
                StatusBrush = "#2ECC71"; // Nice green

                EvaluateDebt();
            }
            else
            {
                // Scenario A - Unregistered Customer
                ResetState();
                IsCustomerFound = false;
                IsFieldsEnabled = true;
                IsJobCardButtonEnabled = true;

                StatusMessage = "❌ Customer Not Found / عميل غير مسجل";
                StatusBrush = "#E74C3C"; // Nice red
            }
        }

        private void EvaluateDebt()
        {
            // Do not disable the job card button anymore, but show the accept cash option if there is debt
            IsJobCardButtonEnabled = true;
            if (TotalDebt > 0)
            {
                IsAcceptCashVisible = true;
            }
            else
            {
                IsAcceptCashVisible = false;
            }
        }

        private void OnRegisterOrOpen(object? parameter)
        {
            string clean = NormalizePhone(SearchPhone);
            if (string.IsNullOrWhiteSpace(clean) || clean.Length < 7)
            {
                MessageBox.Show("Please enter a valid phone number (at least 7-8 digits) / يرجى إدخال رقم هاتف صحيح من 7 أرقام على الأقل.", "Error / خطأ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (!IsCustomerFound)
            {
                // Register customer
                if (string.IsNullOrWhiteSpace(CustomerName))
                {
                    MessageBox.Show("Please enter Customer Name / يرجى إدخال اسم العميل.", "Error / خطأ", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                try
                {
                    string insertQuery = @"
                        INSERT INTO Customers (Name, MobilePhone, WorkPhone, OtherPhone, Address, TotalDebt, CreatedAt)
                        VALUES (@name, @mobile, @work, @other, @address, 0.0, @created);
                        SELECT last_insert_rowid();";

                    var cmdParams = new[]
                    {
                        new SqliteParameter("@name", CustomerName.Trim()),
                        new SqliteParameter("@mobile", SearchPhone),
                        new SqliteParameter("@work", string.IsNullOrEmpty(WorkPhone) ? DBNull.Value : (object)WorkPhone),
                        new SqliteParameter("@other", string.IsNullOrEmpty(OtherPhone) ? DBNull.Value : (object)OtherPhone),
                        new SqliteParameter("@address", string.IsNullOrEmpty(CustomerAddress) ? DBNull.Value : (object)CustomerAddress),
                        new SqliteParameter("@created", DateTime.UtcNow.ToString("o"))
                    };

                    object? result = DatabaseHelper.ExecuteScalar(insertQuery, cmdParams);
                    if (result != null)
                    {
                        _currentCustomerId = Convert.ToInt32(result);
                        IsCustomerFound = true;
                        IsFieldsEnabled = false;
                        StatusMessage = "✔️ Registered Customer / عميل مسجل";
                        StatusBrush = "#2ECC71";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Customer registration failed / فشل تسجيل العميل: {ex.Message}", "Error / خطأ", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            // نقوم بتمرير النافذة الحالية التي استقبلناها من الـ UI
            OpenJobCardWindow(parameter as Window);
        }

        private void OnAcceptCash(object? parameter)
        {
            string confirmMsg =
                $"Have you received {TotalDebt} QAR cash now in the register? / " +
                $"هل استلمت مبلغ {TotalDebt} ر.ق نقداً في الخزينة الآن؟";

            var result = MessageBox.Show(
                confirmMsg,
                "Payment Confirmation / تأكيد السداد",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    double debtCleared = TotalDebt;
                    DatabaseHelper.ExecuteTransaction((conn, trans) =>
                    {
                        // 1. Clear customer debt
                        string updateCust = "UPDATE Customers SET TotalDebt = 0.0 WHERE CustomerID = @id;";
                        using (var cmd = new SqliteCommand(updateCust, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@id", _currentCustomerId);
                            cmd.ExecuteNonQuery();
                        }

                        // 2. Insert cash payment receipt
                        string insertPayment = @"
                            INSERT INTO Payments (CustomerID, Amount, PaymentType, ReceivedBy, ReceivedAt, Notes)
                            VALUES (@custID, @amount, 'Cash', @user, @date, 'Debt fully cleared at card entry');";
                        using (var cmd = new SqliteCommand(insertPayment, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@custID", _currentCustomerId);
                            cmd.Parameters.AddWithValue("@amount", debtCleared);
                            cmd.Parameters.AddWithValue("@user", CurrentUsername);
                            cmd.Parameters.AddWithValue("@date", DateTime.UtcNow.ToString("o"));
                            cmd.ExecuteNonQuery();
                        }

                        // 3. Audit trail with millisecond-precision timestamp
                        string insertAudit = @"
                            INSERT INTO AuditTrail (ActionType, EntityType, EntityID, Description, OldValue, NewValue, UserName, Timestamp)
                            VALUES ('DEBT_CLEARED', 'Customer', @custID, 'Outstanding debt fully settled in cash at entry', @oldVal, '0', @user, @time);";
                        using (var cmd = new SqliteCommand(insertAudit, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@custID", _currentCustomerId);
                            cmd.Parameters.AddWithValue("@oldVal", debtCleared.ToString("F2"));
                            cmd.Parameters.AddWithValue("@user", CurrentUsername);
                            cmd.Parameters.AddWithValue("@time", DateTime.UtcNow.ToString("o"));
                            cmd.ExecuteNonQuery();
                        }
                    });

                    TotalDebt = 0.0;
                    IsJobCardButtonEnabled = true;
                    IsAcceptCashVisible = false;

                    MessageBox.Show(
                        "✅ Debt cleared. Job Card is now unlocked / تم سداد المديونية بنجاح وفتح البطاقة.",
                        "Success / نجاح", MessageBoxButton.OK, MessageBoxImage.Information);

                    // نقوم بتمرير النافذة الحالية التي استقبلناها من الـ UI
                    OpenJobCardWindow(parameter as Window);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Settle payment failed / فشل معالجة السداد: {ex.Message}",
                        "Error / خطأ", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show(
                    $"Please collect the remaining balance: {TotalDebt} QAR before proceeding.\n" +
                    $"يرجى سداد المبلغ المتبقي: {TotalDebt} ر.ق قبل المتابعة.",
                    "Payment Required / يرجى سداد المبلغ",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void OpenJobCardWindow(Window? currentWindow)
        {
            // 1. فتح نافذة كرت العمل للعميل الحالي
            var jobCardWindow = new JobCardWindow(_currentCustomerId);
            jobCardWindow.Show();

            // 2. إغلاق نافذة البحث الحالية فوراً حتى لا تبقى في الخلفية
            currentWindow?.Close();
        }
    }
}
