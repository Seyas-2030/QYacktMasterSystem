using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using QYachtMaster.MVVM;
using QYachtMaster.Database;
using QYachtMaster.Models;
using QYachtMaster.Inventory;
using QYachtMaster.Views;

namespace QYachtMaster.ViewModels
{
    public class JobCardViewModel : ViewModelBase
    {
        private readonly int _customerId;
        private string _customerName  = string.Empty;
        private string _customerAddress = string.Empty;
        private string _mobilePhone   = string.Empty;
        private string _workPhone     = string.Empty;
        private string _otherPhone    = string.Empty;

        // Vehicle details
        private string _regNo     = string.Empty;
        private string _year      = string.Empty;
        private string _make      = string.Empty;
        private string _model     = string.Empty;
        private string _color     = string.Empty;
        private string _odometer  = string.Empty;

        // Job Details
        private string _jobCardNo      = string.Empty;
        private string _arrivalTime    = string.Empty;
        private string _jobDate        = string.Empty;
        private string _pickupTime     = string.Empty;
        private string _pickupPeriod   = "PM";
        private string _dropoffTime    = string.Empty;
        private string _dropoffPeriod  = "AM";

        // ----------------------------------------------------------------
        // Stock / Cascading Dropdown state
        // ----------------------------------------------------------------
        private string? _stockFilePath;
        private string? _selectedDepartment;
        private TechPartOption? _selectedPartFromDropdown;

        /// <summary>Level-1 dropdown: department names from Excel sheets + manual items.</summary>
        public ObservableCollection<string> Departments { get; } = new();

        /// <summary>Level-2 dropdown: filtered parts for the selected department.</summary>
        public ObservableCollection<TechPartOption> FilteredParts { get; } = new();

        /// <summary>Selected items (chips) — transfers to invoice on save.</summary>
        public ObservableCollection<SelectedPartItem> SelectedParts { get; } = new();

        public string? SelectedDepartment
        {
            get => _selectedDepartment;
            set
            {
                if (SetProperty(ref _selectedDepartment, value))
                {
                    LoadPartsForDepartment(value);
                }
            }
        }

        public TechPartOption? SelectedPartFromDropdown
        {
            get => _selectedPartFromDropdown;
            set
            {
                if (SetProperty(ref _selectedPartFromDropdown, value) && value != null)
                {
                    AddSelectedPart(value);
                    _selectedPartFromDropdown = null;
                    OnPropertyChanged(nameof(SelectedPartFromDropdown));
                }
            }
        }

        // The 8 Repair Action Lines
        public ObservableCollection<RepairLineItem> RepairLines { get; } = new();

        // Read-only display props
        public string CustomerName    => _customerName;
        public string CustomerAddress => _customerAddress;
        public string MobilePhone     => _mobilePhone;
        public string WorkPhone       => _workPhone;
        public string OtherPhone      => _otherPhone;

        public string JobCardNo
        {
            get => _jobCardNo;
            set => SetProperty(ref _jobCardNo, value);
        }

        public string RegNo
        {
            get => _regNo;
            set => SetProperty(ref _regNo, value);
        }

        public string Year
        {
            get => _year;
            set => SetProperty(ref _year, value);
        }

        public string Make
        {
            get => _make;
            set => SetProperty(ref _make, value);
        }

        public string Model
        {
            get => _model;
            set => SetProperty(ref _model, value);
        }

        public string Color
        {
            get => _color;
            set => SetProperty(ref _color, value);
        }

        public string Odometer
        {
            get => _odometer;
            set => SetProperty(ref _odometer, value);
        }

        public string ArrivalTime
        {
            get => _arrivalTime;
            set => SetProperty(ref _arrivalTime, value);
        }

        public string JobDate
        {
            get => _jobDate;
            set => SetProperty(ref _jobDate, value);
        }

        public string PickupTime
        {
            get => _pickupTime;
            set => SetProperty(ref _pickupTime, value);
        }

        public string PickupPeriod
        {
            get => _pickupPeriod;
            set => SetProperty(ref _pickupPeriod, value);
        }

        public string DropoffTime
        {
            get => _dropoffTime;
            set => SetProperty(ref _dropoffTime, value);
        }

        public string DropoffPeriod
        {
            get => _dropoffPeriod;
            set => SetProperty(ref _dropoffPeriod, value);
        }

        public ICommand SaveAndProceedCommand { get; }
        public ICommand RemovePartCommand     { get; }
        public ICommand ChangeStockFileCommand { get; }

        public JobCardViewModel(int customerId)
        {
            _customerId = customerId;

            SaveAndProceedCommand  = new RelayCommand(OnSaveAndProceed);
            RemovePartCommand      = new RelayCommand(OnRemovePart);
            ChangeStockFileCommand = new RelayCommand(_ => PromptForStockFile(required: false));

            LoadCustomerData();
            GenerateJobCardNumber();

            for (int i = 1; i <= 8; i++)
                RepairLines.Add(new RepairLineItem { LineNumber = i });

            JobDate     = DateTime.Now.ToString("yyyy-MM-dd");
            ArrivalTime = DateTime.Now.ToString("hh:mm tt");

            InitializeStockAndDepartments();
        }

        // ----------------------------------------------------------------
        // Stock file bootstrap
        // ----------------------------------------------------------------
        private void InitializeStockAndDepartments()
        {
            // Retrieve cached file path from DB settings
            _stockFilePath = DatabaseInitializer.GetSetting("StockFilePath");

            if (string.IsNullOrEmpty(_stockFilePath) || !File.Exists(_stockFilePath))
            {
                // Prompt user to locate the file
                PromptForStockFile(required: true);
            }
            else
            {
                LoadDepartmentsFromExcel();
            }
        }

        private void PromptForStockFile(bool required)
        {
            var dlg = new OpenFileDialog
            {
                Title  = "حدد ملف المخزون / Locate Stock File (stock2026.xlsx)",
                Filter = "Excel Files|*.xlsx",
                FileName = "stock2026.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                _stockFilePath = dlg.FileName;
                DatabaseInitializer.SetSetting("StockFilePath", _stockFilePath);
                LoadDepartmentsFromExcel();
            }
            else if (required)
            {
                MessageBox.Show(
                    "ملف المخزون مطلوب لتشغيل القوائم المنسدلة. يمكنك تغييره لاحقاً.\n" +
                    "The stock file is required for the dropdown menus. You can change it later.",
                    "تنبيه / Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                // Load manual-only departments so the app still works
                LoadManualDepartments();
            }
        }

        private void LoadDepartmentsFromExcel()
        {
            Departments.Clear();
            try
            {
                var sheetNames = PartsSyncEngine.GetSheetNames(_stockFilePath!);
                foreach (var name in sheetNames)
                    Departments.Add(name);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"تعذّر فتح ملف المخزون: {ex.Message}\nCould not open stock file: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Always append manual items at the end
            LoadManualDepartments();
        }

        private void LoadManualDepartments()
        {
            if (!Departments.Contains("Electrical Problems / مشكلة كهربائية"))
                Departments.Add("Electrical Problems / مشكلة كهربائية");

            if (!Departments.Contains("Computer Check / فحص كمبيوتر"))
                Departments.Add("Computer Check / فحص كمبيوتر");
        }

        private void LoadPartsForDepartment(string? department)
        {
            FilteredParts.Clear();
            if (string.IsNullOrEmpty(department)) return;

            // Manual departments — no parts list
            if (department.StartsWith("Electrical") || department.StartsWith("Computer"))
            {
                FilteredParts.Add(new TechPartOption
                {
                    PartNo    = department,
                    Model     = "",
                    SalePrice = 0.0,
                    CostPrice = 0.0,
                    DisplayLabel = department
                });
                return;
            }

            if (string.IsNullOrEmpty(_stockFilePath) || !File.Exists(_stockFilePath)) return;

            try
            {
                var parts = PartsSyncEngine.GetPartsForSheet(_stockFilePath, department);
                foreach (var p in parts)
                {
                    // Build a rich label:  [BoxLabel] PartNo — Model — Price
                    string boxPrefix = string.IsNullOrEmpty(p.BoxLabel) ? "" : $"[{p.BoxLabel}] ";
                    string priceStr  = p.SalePrice > 0 ? $" — {p.SalePrice:F2} QAR" : "";
                    string modelStr  = string.IsNullOrEmpty(p.Model) ? "" : $" — {p.Model}";

                    FilteredParts.Add(new TechPartOption
                    {
                        PartNo    = p.PartNo,
                        Model     = p.Model,
                        SalePrice = p.SalePrice,
                        CostPrice = p.CostPrice,
                        DisplayLabel = $"{boxPrefix}{p.PartNo}{modelStr}{priceStr}"
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطأ في تحميل القطع: {ex.Message}\nError loading parts: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddSelectedPart(TechPartOption option)
        {
            // Prevent duplicate PartNo
            if (SelectedParts.Any(p => p.PartNo == option.PartNo && p.Department == SelectedDepartment))
                return;

            SelectedParts.Add(new SelectedPartItem
            {
                PartNo     = option.PartNo,
                Department = SelectedDepartment ?? string.Empty,
                Model      = option.Model,
                SalePrice  = option.SalePrice,
                CostPrice  = option.CostPrice,
                DisplayLabel = $"[{SelectedDepartment}] {option.DisplayLabel}"
            });
        }

        private void OnRemovePart(object? parameter)
        {
            if (parameter is SelectedPartItem item)
                SelectedParts.Remove(item);
        }

        // ----------------------------------------------------------------
        // Customer & card helpers
        // ----------------------------------------------------------------
        private void LoadCustomerData()
        {
            string query = "SELECT * FROM Customers WHERE CustomerID = @id;";
            var p = new SqliteParameter("@id", _customerId);
            var customer = DatabaseHelper.ExecuteSingleQuery(query, reader => new Customer
            {
                Name       = reader["Name"].ToString()!,
                MobilePhone = reader["MobilePhone"].ToString()!,
                WorkPhone  = reader["WorkPhone"]?.ToString(),
                OtherPhone = reader["OtherPhone"]?.ToString(),
                Address    = reader["Address"]?.ToString()
            }, p);

            if (customer != null)
            {
                _customerName    = customer.Name;
                _customerAddress = customer.Address ?? string.Empty;
                _mobilePhone     = customer.MobilePhone;
                _workPhone       = customer.WorkPhone ?? string.Empty;
                _otherPhone      = customer.OtherPhone ?? string.Empty;

                OnPropertyChanged(nameof(CustomerName));
                OnPropertyChanged(nameof(CustomerAddress));
                OnPropertyChanged(nameof(MobilePhone));
                OnPropertyChanged(nameof(WorkPhone));
                OnPropertyChanged(nameof(OtherPhone));

                // Retrieve vehicle details for existing customer if available
                string vehQuery = "SELECT * FROM Vehicles WHERE CustomerID = @id ORDER BY VehicleID DESC LIMIT 1;";
                var vp = new SqliteParameter("@id", _customerId);
                var vehicle = DatabaseHelper.ExecuteSingleQuery(vehQuery, reader => new Vehicle
                {
                    RegNo = reader["RegNo"]?.ToString() ?? string.Empty,
                    Year = reader["Year"]?.ToString() ?? string.Empty,
                    Make = reader["Make"]?.ToString() ?? string.Empty,
                    Model = reader["Model"]?.ToString() ?? string.Empty,
                    Color = reader["Color"]?.ToString() ?? string.Empty,
                    LastOdometer = reader["LastOdometer"]?.ToString() ?? string.Empty
                }, vp);

                if (vehicle != null)
                {
                    _regNo = vehicle.RegNo ?? string.Empty;
                    _year = vehicle.Year ?? string.Empty;
                    _make = vehicle.Make ?? string.Empty;
                    _model = vehicle.Model ?? string.Empty;
                    _color = vehicle.Color ?? string.Empty;
                    _odometer = vehicle.LastOdometer ?? string.Empty;

                    OnPropertyChanged(nameof(RegNo));
                    OnPropertyChanged(nameof(Year));
                    OnPropertyChanged(nameof(Make));
                    OnPropertyChanged(nameof(Model));
                    OnPropertyChanged(nameof(Color));
                    OnPropertyChanged(nameof(Odometer));
                }
            }
        }

        private void GenerateJobCardNumber()
        {
            int currentYear = DateTime.Now.Year;
            string prefix   = $"JC-{currentYear}-";

            string query   = $"SELECT JobCardNo FROM JobCards WHERE JobCardNo LIKE '{prefix}%' ORDER BY JobCardID DESC LIMIT 1;";
            string? lastNo = DatabaseHelper.ExecuteScalar(query) as string;

            int sequence = 1;
            if (!string.IsNullOrEmpty(lastNo))
            {
                var parts = lastNo.Split('-');
                if (parts.Length == 3 && int.TryParse(parts[2], out int parsedSeq))
                    sequence = parsedSeq + 1;
            }

            JobCardNo = $"{prefix}{sequence:D4}";
        }

        // ----------------------------------------------------------------
        // Save
        // ----------------------------------------------------------------
        private void OnSaveAndProceed(object? parameter)
        {
            if (string.IsNullOrWhiteSpace(RegNo))
            {
                MessageBox.Show("يرجى إدخال رقم تسجيل السيارة / Please enter Vehicle Reg No.",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                int vehicleId = 0;
                int jobCardId = 0;

                DatabaseHelper.ExecuteTransaction((conn, trans) =>
                {
                    // Upsert vehicle
                    string queryVehicle = "SELECT VehicleID FROM Vehicles WHERE CustomerID = @custID AND RegNo = @reg;";
                    using (var cmd = new SqliteCommand(queryVehicle, conn, trans))
                    {
                        cmd.Parameters.AddWithValue("@custID", _customerId);
                        cmd.Parameters.AddWithValue("@reg", RegNo.Trim());
                        var result = cmd.ExecuteScalar();
                        if (result != null)
                        {
                            vehicleId = Convert.ToInt32(result);
                            string updateVeh = "UPDATE Vehicles SET Year=@y, Make=@mk, Model=@md, Color=@col, LastOdometer=@odo WHERE VehicleID=@vid;";
                            using var updateCmd = new SqliteCommand(updateVeh, conn, trans);
                            updateCmd.Parameters.AddWithValue("@y",   Year    ?? string.Empty);
                            updateCmd.Parameters.AddWithValue("@mk",  Make    ?? string.Empty);
                            updateCmd.Parameters.AddWithValue("@md",  Model   ?? string.Empty);
                            updateCmd.Parameters.AddWithValue("@col", Color   ?? string.Empty);
                            updateCmd.Parameters.AddWithValue("@odo", Odometer ?? string.Empty);
                            updateCmd.Parameters.AddWithValue("@vid", vehicleId);
                            updateCmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string insertVeh = @"
                                INSERT INTO Vehicles (CustomerID, RegNo, Year, Make, Model, Color, LastOdometer)
                                VALUES (@custID, @reg, @y, @mk, @md, @col, @odo);
                                SELECT last_insert_rowid();";
                            using var insertCmd = new SqliteCommand(insertVeh, conn, trans);
                            insertCmd.Parameters.AddWithValue("@custID", _customerId);
                            insertCmd.Parameters.AddWithValue("@reg",    RegNo.Trim());
                            insertCmd.Parameters.AddWithValue("@y",      Year     ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@mk",     Make     ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@md",     Model    ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@col",    Color    ?? string.Empty);
                            insertCmd.Parameters.AddWithValue("@odo",    Odometer ?? string.Empty);
                            vehicleId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        }
                    }

                    // Create Job Card
                    string insertJob = @"
                        INSERT INTO JobCards (JobCardNo, CustomerID, VehicleID, ArrivalTime, JobDate,
                            PickupTime, PickupPeriod, DropoffTime, DropoffPeriod, OdometerAtVisit,
                            Status, CreatedAt, CreatedBy)
                        VALUES (@jcNo, @custID, @vehID, @arr, @jobDate, @pickTime, @pickPeriod,
                            @dropTime, @dropPeriod, @odo, 'In Progress', @created, 'admin');
                        SELECT last_insert_rowid();";

                    using (var cmd = new SqliteCommand(insertJob, conn, trans))
                    {
                        cmd.Parameters.AddWithValue("@jcNo",       JobCardNo);
                        cmd.Parameters.AddWithValue("@custID",     _customerId);
                        cmd.Parameters.AddWithValue("@vehID",      vehicleId);
                        cmd.Parameters.AddWithValue("@arr",        ArrivalTime    ?? string.Empty);
                        cmd.Parameters.AddWithValue("@jobDate",    JobDate);
                        cmd.Parameters.AddWithValue("@pickTime",   PickupTime     ?? string.Empty);
                        cmd.Parameters.AddWithValue("@pickPeriod", PickupPeriod   ?? string.Empty);
                        cmd.Parameters.AddWithValue("@dropTime",   DropoffTime    ?? string.Empty);
                        cmd.Parameters.AddWithValue("@dropPeriod", DropoffPeriod  ?? string.Empty);
                        cmd.Parameters.AddWithValue("@odo",        Odometer       ?? string.Empty);
                        cmd.Parameters.AddWithValue("@created",    DateTime.UtcNow.ToString("o"));
                        jobCardId = Convert.ToInt32(cmd.ExecuteScalar());
                    }

                    // Save 8 Repair Lines
                    foreach (var line in RepairLines)
                    {
                        if (!string.IsNullOrEmpty(line.ServiceDescription) || !string.IsNullOrEmpty(line.TechnicianName))
                        {
                            string insertLine = @"
                                INSERT INTO JobCardRepairLines (JobCardID, LineNumber, ServiceDescription, TechnicianName)
                                VALUES (@jcID, @lineNum, @desc, @tech);";
                            using var cmd = new SqliteCommand(insertLine, conn, trans);
                            cmd.Parameters.AddWithValue("@jcID",    jobCardId);
                            cmd.Parameters.AddWithValue("@lineNum", line.LineNumber);
                            cmd.Parameters.AddWithValue("@desc",    line.ServiceDescription ?? string.Empty);
                            cmd.Parameters.AddWithValue("@tech",    line.TechnicianName     ?? string.Empty);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    // Save Selected Parts as JobCardTechComments (ServiceKey = PartNo)
                    foreach (var part in SelectedParts)
                    {
                        string insertComment = @"
                            INSERT INTO JobCardTechComments (JobCardID, ServiceKey, IsSelected, Notes)
                            VALUES (@jcID, @key, 1, @notes);";
                        using var cmd = new SqliteCommand(insertComment, conn, trans);
                        cmd.Parameters.AddWithValue("@jcID",  jobCardId);
                        cmd.Parameters.AddWithValue("@key",   part.PartNo);
                        cmd.Parameters.AddWithValue("@notes", $"{part.Department}|{part.Model}|{part.SalePrice}|{part.CostPrice}");
                        cmd.ExecuteNonQuery();
                    }
                });

                // Open Invoice Window
                var invoiceWindow = new InvoiceWindow(jobCardId, _stockFilePath);
                invoiceWindow.Show();

                if (parameter is Window window)
                    window.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل حفظ بطاقة العمل / Failed to save Job Card: {ex.Message}",
                    "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // ----------------------------------------------------------------
    // Supporting types
    // ----------------------------------------------------------------

    public class RepairLineItem
    {
        public int    LineNumber          { get; set; }
        public string ServiceDescription { get; set; } = string.Empty;
        public string TechnicianName     { get; set; } = string.Empty;
    }

    public class TechPartOption
    {
        public string PartNo       { get; set; } = string.Empty;
        public string Model        { get; set; } = string.Empty;
        public double SalePrice    { get; set; }
        public double CostPrice    { get; set; }
        public string DisplayLabel { get; set; } = string.Empty;
    }

    public class SelectedPartItem
    {
        public string PartNo       { get; set; } = string.Empty;
        public string Department   { get; set; } = string.Empty;
        public string Model        { get; set; } = string.Empty;
        public double SalePrice    { get; set; }
        public double CostPrice    { get; set; }
        public string DisplayLabel { get; set; } = string.Empty;
    }
}
