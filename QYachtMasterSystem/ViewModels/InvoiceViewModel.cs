using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Microsoft.Data.Sqlite;
using QYachtMaster.MVVM;
using QYachtMaster.Database;
using QYachtMaster.Models;
using QYachtMaster.Views;
using QYachtMaster.Printing;
using ClosedXML.Excel;

namespace QYachtMaster.ViewModels
{
    public class InvoiceViewModel : ViewModelBase
    {
        private readonly int _jobCardId;
        private readonly string? _stockFilePath;   // Path to stock2026.xlsx for write-back
        private int _customerId;
        
        private string _invoiceNo = string.Empty;
        private string _customerName = string.Empty;
        private string _mobilePhone = string.Empty;
        private string _vehicleDetails = string.Empty;
        private string _invoiceDate = string.Empty;
        
        private double _subTotal = 0.0;
        private double _totalDiscounts = 0.0;
        private double _laborTotal = 0.0;
        private double _grandTotal = 0.0;
        private double _paidAmount = 0.0;
        private string _paymentMethod = "Cash"; // Default

        private bool _isLocked = false;
        private string _lockStatusText = string.Empty;
        private string _stockWarningMessage = string.Empty;
        private bool _hasStockWarning = false;
        
        private double _previousDebt = 0.0;
        private double _debtPayment = 0.0;

        public double PreviousDebt
        {
            get => _previousDebt;
            set => SetProperty(ref _previousDebt, value);
        }

        public double DebtPayment
        {
            get => _debtPayment;
            set
            {
                if (SetProperty(ref _debtPayment, value))
                {
                    RecalculateTotals();
                }
            }
        }

        public ObservableCollection<InvoiceLineViewModel> InvoiceLines { get; } = new();

        public string InvoiceNo
        {
            get => _invoiceNo;
            set => SetProperty(ref _invoiceNo, value);
        }

        public string CustomerName => _customerName;
        public string MobilePhone
        {
            get => _mobilePhone;
            set => SetProperty(ref _mobilePhone, value);
        }
        public string VehicleDetails => _vehicleDetails;

        public string InvoiceDate
        {
            get => _invoiceDate;
            set => SetProperty(ref _invoiceDate, value);
        }

        public double SubTotal
        {
            get => _subTotal;
            set => SetProperty(ref _subTotal, value);
        }

        public double TotalDiscounts
        {
            get => _totalDiscounts;
            set => SetProperty(ref _totalDiscounts, value);
        }

        public double LaborTotal
        {
            get => _laborTotal;
            set
            {
                if (SetProperty(ref _laborTotal, value))
                {
                    RecalculateTotals();
                }
            }
        }

        public double GrandTotal
        {
            get => _grandTotal;
            set
            {
                if (SetProperty(ref _grandTotal, value))
                {
                    OnPropertyChanged(nameof(BalanceDue));
                }
            }
        }

        public double PaidAmount
        {
            get => _paidAmount;
            set
            {
                if (SetProperty(ref _paidAmount, value))
                {
                    OnPropertyChanged(nameof(BalanceDue));
                }
            }
        }

        public double BalanceDue => Math.Max(0.0, GrandTotal - PaidAmount);

        public string PaymentMethod
        {
            get => _paymentMethod;
            set => SetProperty(ref _paymentMethod, value);
        }

        public bool IsLocked
        {
            get => _isLocked;
            set
            {
                if (SetProperty(ref _isLocked, value))
                {
                    OnPropertyChanged(nameof(IsEditable));
                }
            }
        }

        public bool IsEditable => !IsLocked;

        public string LockStatusText
        {
            get => _lockStatusText;
            set => SetProperty(ref _lockStatusText, value);
        }

        public string StockWarningMessage
        {
            get => _stockWarningMessage;
            set => SetProperty(ref _stockWarningMessage, value);
        }

        public bool HasStockWarning
        {
            get => _hasStockWarning;
            set => SetProperty(ref _hasStockWarning, value);
        }

        public ICommand AddCustomLineCommand { get; }
        public ICommand RemoveLineCommand { get; }
        public ICommand ImportExcelCommand { get; }
        public ICommand SaveAndPrintCommand { get; }

        public InvoiceViewModel(int jobCardId, string? stockFilePath = null)
        {
            _jobCardId     = jobCardId;
            _stockFilePath = stockFilePath;
            
            AddCustomLineCommand = new RelayCommand(OnAddCustomLine);
            RemoveLineCommand = new RelayCommand(OnRemoveLine);
            ImportExcelCommand = new RelayCommand(OnImportExcel);
            SaveAndPrintCommand = new RelayCommand(OnSaveAndPrint);

            InvoiceDate = DateTime.Now.ToString("yyyy-MM-dd");

            LoadJobCardData();
            GenerateInvoiceNumber();
            PopulateLinesFromTechComments();
        }

        private void LoadJobCardData()
        {
            string query = @"
                SELECT j.CustomerID, c.Name, c.MobilePhone, c.TotalDebt, v.RegNo, v.Make, v.Model, v.Year
                FROM JobCards j
                JOIN Customers c ON j.CustomerID = c.CustomerID
                JOIN Vehicles v ON j.VehicleID = v.VehicleID
                WHERE j.JobCardID = @jcID;";

            var p = new SqliteParameter("@jcID", _jobCardId);
            DatabaseHelper.ExecuteSingleQuery(query, reader =>
            {
                _customerId = Convert.ToInt32(reader["CustomerID"]);
                _customerName = reader["Name"].ToString()!;
                _mobilePhone = reader["MobilePhone"].ToString()!;
                _previousDebt = Convert.ToDouble(reader["TotalDebt"]);
                
                string reg = reader["RegNo"]?.ToString() ?? string.Empty;
                string mk = reader["Make"]?.ToString() ?? string.Empty;
                string md = reader["Model"]?.ToString() ?? string.Empty;
                string yr = reader["Year"]?.ToString() ?? string.Empty;
                _vehicleDetails = $"{reg} - {mk} {md} ({yr})";

                return 0;
            }, p);
        }

        private void GenerateInvoiceNumber()
        {
            // Invoice number starts at 1000 on first install, increments by 1 thereafter.
            int currentYear = DateTime.Now.Year;
            string prefix   = $"INV-{currentYear}-";

            string query   = $"SELECT InvoiceNo FROM Invoices WHERE InvoiceNo LIKE '{prefix}%' ORDER BY InvoiceID DESC LIMIT 1;";
            string? lastNo = DatabaseHelper.ExecuteScalar(query) as string;

            int sequence = 1000; // Always start at 1000 on a fresh install
            if (!string.IsNullOrEmpty(lastNo))
            {
                var parts = lastNo.Split('-');
                if (parts.Length == 3 && int.TryParse(parts[2], out int parsedSeq))
                    sequence = parsedSeq + 1;  // Increment by 1
            }

            InvoiceNo = $"{prefix}{sequence}";
        }

        private void PopulateLinesFromTechComments()
        {
            // Retrieve all selected tech comments (ServiceKey = PartNo, Notes = dept|model|salePrice|costPrice)
            string query = @"
                SELECT ServiceKey, Notes FROM JobCardTechComments 
                WHERE JobCardID = @jcID AND IsSelected = 1;";
            
            var p = new SqliteParameter("@jcID", _jobCardId);

            // Fetch to in-memory list first to avoid nested database readers on the same connection
            var selectedComments = DatabaseHelper.ExecuteQuery(query, reader => new
            {
                ServiceKey = reader["ServiceKey"]?.ToString() ?? string.Empty,
                Notes = reader["Notes"]?.ToString() ?? string.Empty
            }, p);

            int order = 1;
            foreach (var item in selectedComments)
            {
                string partNo = item.ServiceKey;
                string notes  = item.Notes;

                double salePrice = 0.0;
                double costPrice = 0.0;
                string desc      = partNo;
                double stockQty  = 99.0;

                // Parse new format: department|model|salePrice|costPrice
                var noteParts = notes.Split('|');
                if (noteParts.Length >= 4)
                {
                    string dept  = noteParts[0];
                    string model = noteParts[1];
                    double.TryParse(noteParts[2], out salePrice);
                    double.TryParse(noteParts[3], out costPrice);
                    desc = string.IsNullOrEmpty(model) ? $"[{dept}] {partNo}" : $"[{dept}] {partNo} — {model}";

                    // Check stock qty from DB for low-stock alert
                    string stockQuery = "SELECT QtyInStock FROM Parts WHERE PartNo = @pn LIMIT 1;";
                    var sp = new SqliteParameter("@pn", partNo);
                    DatabaseHelper.ExecuteSingleQuery(stockQuery, r =>
                    {
                        stockQty = Convert.ToDouble(r["QtyInStock"]);
                        return 0;
                    }, sp);
                }
                else
                {
                    // Fallback: legacy static key format — look up from Parts by PartNo or Category
                    string partQuery = "SELECT SalePrice, CostPrice, QtyInStock FROM Parts WHERE PartNo = @pn LIMIT 1;";
                    var cp = new SqliteParameter("@pn", partNo);
                    DatabaseHelper.ExecuteSingleQuery(partQuery, r =>
                    {
                        salePrice = Convert.ToDouble(r["SalePrice"]);
                        costPrice = Convert.ToDouble(r["CostPrice"]);
                        stockQty  = Convert.ToDouble(r["QtyInStock"]);
                        return 0;
                    }, cp);
                }

                var line = new InvoiceLineViewModel(this)
                {
                    Index            = order++,
                    Description      = desc,
                    Qty              = 1.0,
                    UnitPrice        = salePrice,
                    UnitCostPrice    = costPrice,
                    Discount         = 0.0,
                    SourceServiceKey = partNo
                };

                InvoiceLines.Add(line);


                // Low Stock Alert: warn when stock is critically low (≤1) or already depleted
                if (stockQty <= 0.0)
                {
                    HasStockWarning     = true;
                    StockWarningMessage += $"🚫 [{desc}]: نفذت القطعة من المخزن / OUT OF STOCK — please restock immediately\n";
                }
                else if (stockQty <= 1.0)
                {
                    HasStockWarning     = true;
                    StockWarningMessage += $"⚠️ [{desc}]: قطعة واحدة فقط متبقية — يرجى تعبئة المخزن / Only 1 unit left — please restock\n";
                }

            }

            RecalculateTotals();
        }


        public void RecalculateTotals()
        {
            double sub = 0.0;
            double disc = 0.0;

            for (int i = 0; i < InvoiceLines.Count; i++)
            {
                var line = InvoiceLines[i];
                line.Index = i + 1; // Update index order on-the-fly
                sub += line.QtySubTotal;
                disc += line.Discount;
            }

            SubTotal = sub;
            TotalDiscounts = disc;
            GrandTotal = Math.Max(0.0, SubTotal - TotalDiscounts + LaborTotal + DebtPayment);
        }

        private void OnAddCustomLine(object? parameter)
        {
            var line = new InvoiceLineViewModel(this)
            {
                Index = InvoiceLines.Count + 1,
                Description = "خدمة مخصصة / Custom Service",
                Qty = 1.0,
                UnitPrice = 0.0,
                Discount = 0.0
            };
            InvoiceLines.Add(line);
            RecalculateTotals();
        }

        private void OnRemoveLine(object? parameter)
        {
            if (parameter is InvoiceLineViewModel line)
            {
                InvoiceLines.Remove(line);
                RecalculateTotals();
            }
        }

        private void OnImportExcel(object? parameter)
        {
            // Prompt mapping for dynamic excel import (Module 4)
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select Excel File to Import Invoice Lines / اختر ملف إكسل لاستيراد البنود"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                try
                {
                    // Open column mapping view modal
                    var excelData = new DataTable();
                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var worksheet = workbook.Worksheets.First();
                        var firstRow = worksheet.FirstRowUsed();
                        if (firstRow == null) return;

                        foreach (var cell in firstRow.Cells())
                        {
                            excelData.Columns.Add(cell.GetString());
                        }

                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var newRow = excelData.NewRow();
                            for (int i = 0; i < excelData.Columns.Count; i++)
                            {
                                newRow[i] = row.Cell(i + 1).GetString();
                            }
                            excelData.Rows.Add(newRow);
                        }
                    }

                    // Prompt mapper dialog (simplifying mapping options inline for UX or via direct selection dialog)
                    var mapper = new ExcelColumnMapperWindow(excelData);
                    if (mapper.ShowDialog() == true)
                    {
                        string partNoCol = mapper.PartNoColumn;
                        string qtyCol = mapper.QtyColumn;
                        string salePriceCol = mapper.SalePriceColumn;
                        string descCol = mapper.DescriptionColumn;

                        foreach (DataRow row in excelData.Rows)
                        {
                            string partNo = row[partNoCol]?.ToString() ?? string.Empty;
                            string desc = string.IsNullOrEmpty(descCol) ? partNo : (row[descCol]?.ToString() ?? partNo);
                            double qty = 1.0;
                            double unitPrice = 0.0;

                            if (!string.IsNullOrEmpty(qtyCol) && double.TryParse(CleanQtyText(row[qtyCol]?.ToString() ?? "0"), out double q))
                            {
                                qty = q;
                            }
                            if (!string.IsNullOrEmpty(salePriceCol) && double.TryParse(row[salePriceCol]?.ToString() ?? "0", out double p))
                            {
                                unitPrice = p;
                            }

                            // Add to invoice lines
                            var line = new InvoiceLineViewModel(this)
                            {
                                Index = InvoiceLines.Count + 1,
                                Description = desc,
                                Qty = qty,
                                UnitPrice = unitPrice,
                                Discount = 0.0
                            };
                            InvoiceLines.Add(line);
                        }

                        RecalculateTotals();
                        MessageBox.Show("تم استيراد بنود الفاتورة بنجاح / Invoice lines imported successfully.", "نجاح / Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"فشل استيراد البنود / Failed to import excel lines: {ex.Message}", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string CleanQtyText(string raw)
        {
            var cleaned = Regex.Replace(raw ?? "0", @"[^0-9\-\.]", "");
            return string.IsNullOrEmpty(cleaned) ? "0" : cleaned;
        }

        private void OnSaveAndPrint(object? parameter)
        {
            // 1. Mandatory Interlocking Confirmation Modal
            string msg = "Are you sure? Invoice will be locked after printing. / هل أنت متأكد من طباعة الفاتورة؟ لن يمكن التعديل بعد الطباعة.";
            var confirmResult = MessageBox.Show(msg, "تأكيد الطباعة والقفل / Print & Lock Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            
            if (confirmResult == MessageBoxResult.Yes)
            {
                try
                {
                    int invoiceId = 0;
                    
                    DatabaseHelper.ExecuteTransaction((conn, trans) =>
                    {
                        // 2. Insert Invoice
                        string insertInvoice = @"
                            INSERT INTO Invoices (InvoiceNo, JobCardID, CustomerID, IsLocked, LockedAt, LockedBy, SubTotal, LaborTotal, GrandTotal, PaidAmount, RemainingDebt, PaymentMethod, CreatedAt)
                            VALUES (@invNo, @jcID, @custID, 1, @now, 'admin', @sub, @labor, @grand, @paid, @remain, @method, @now);
                            SELECT last_insert_rowid();";

                        using (var cmd = new SqliteCommand(insertInvoice, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@invNo", InvoiceNo);
                            cmd.Parameters.AddWithValue("@jcID", _jobCardId);
                            cmd.Parameters.AddWithValue("@custID", _customerId);
                            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("o"));
                            cmd.Parameters.AddWithValue("@sub", SubTotal);
                            cmd.Parameters.AddWithValue("@labor", LaborTotal);
                            cmd.Parameters.AddWithValue("@grand", GrandTotal);
                            cmd.Parameters.AddWithValue("@paid", PaidAmount);
                            cmd.Parameters.AddWithValue("@remain", BalanceDue);
                            cmd.Parameters.AddWithValue("@method", PaymentMethod);

                            invoiceId = Convert.ToInt32(cmd.ExecuteScalar());
                        }

                        // 3. Insert InvoiceLines
                        foreach (var line in InvoiceLines)
                        {
                            string insertLine = @"
                                INSERT INTO InvoiceLines (InvoiceID, Description, Qty, UnitPrice, UnitCostPrice, DiscountAmount, DiscountNote, FinalLineTotal, IsManualOverride, SourceServiceKey, Warranty, SortOrder)
                                VALUES (@invID, @desc, @qty, @price, @cost, @disc, @discNote, @finalTotal, @override, @key, @warranty, @sort);";
                            
                            using (var cmd = new SqliteCommand(insertLine, conn, trans))
                            {
                                cmd.Parameters.AddWithValue("@invID",      invoiceId);
                                cmd.Parameters.AddWithValue("@desc",       line.Description);
                                cmd.Parameters.AddWithValue("@qty",        line.Qty);
                                cmd.Parameters.AddWithValue("@price",      line.UnitPrice);
                                cmd.Parameters.AddWithValue("@cost",       line.UnitCostPrice);
                                cmd.Parameters.AddWithValue("@disc",       line.Discount);
                                cmd.Parameters.AddWithValue("@discNote",   line.DiscountNotes   ?? string.Empty);
                                cmd.Parameters.AddWithValue("@finalTotal", line.FinalLineTotal);
                                cmd.Parameters.AddWithValue("@override",   line.IsManualOverride ? 1 : 0);
                                cmd.Parameters.AddWithValue("@key",        line.SourceServiceKey ?? string.Empty);
                                cmd.Parameters.AddWithValue("@warranty",   line.Warranty         ?? string.Empty);
                                cmd.Parameters.AddWithValue("@sort",       line.Index);

                                cmd.ExecuteNonQuery();
                            }

                            // 4. Warehouse Stock Deductions in DB
                            if (!string.IsNullOrEmpty(line.SourceServiceKey))
                            {
                                // Deduct from Parts table using PartNo (SourceServiceKey now stores PartNo directly)
                                string deductStock = "UPDATE Parts SET QtyInStock = MAX(0, QtyInStock - @qty) WHERE PartNo = @partNo;";
                                using (var cmd = new SqliteCommand(deductStock, conn, trans))
                                {
                                    cmd.Parameters.AddWithValue("@qty",    line.Qty);
                                    cmd.Parameters.AddWithValue("@partNo", line.SourceServiceKey);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        // 5. Update Customer TotalDebt (New Debt = Previous Debt - Debt Payment + Current Invoice Balance Due)
                        double newTotalDebt = Math.Max(0.0, _previousDebt - DebtPayment + BalanceDue);
                        string updateDebt = "UPDATE Customers SET TotalDebt = @newDebt WHERE CustomerID = @custID;";
                        using (var cmd = new SqliteCommand(updateDebt, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@newDebt", newTotalDebt);
                            cmd.Parameters.AddWithValue("@custID", _customerId);
                            cmd.ExecuteNonQuery();
                        }

                        // 5b. Log Payments records for ledger audit
                        if (PaidAmount > 0)
                        {
                            string insertPay = @"
                                INSERT INTO Payments (CustomerID, InvoiceID, Amount, PaymentType, ReceivedBy, ReceivedAt)
                                VALUES (@custID, @invID, @amt, @type, 'admin', @now);";
                            using (var cmd = new SqliteCommand(insertPay, conn, trans))
                            {
                                cmd.Parameters.AddWithValue("@custID", _customerId);
                                cmd.Parameters.AddWithValue("@invID", invoiceId);
                                cmd.Parameters.AddWithValue("@amt", PaidAmount);
                                cmd.Parameters.AddWithValue("@type", PaymentMethod ?? "Cash");
                                cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("o"));
                                cmd.ExecuteNonQuery();
                            }
                        }

                        if (DebtPayment > 0)
                        {
                            string insertDebtPay = @"
                                INSERT INTO Payments (CustomerID, InvoiceID, Amount, PaymentType, ReceivedBy, ReceivedAt)
                                VALUES (@custID, @invID, @amt, 'Debt Settlement', 'admin', @now);";
                            using (var cmd = new SqliteCommand(insertDebtPay, conn, trans))
                            {
                                cmd.Parameters.AddWithValue("@custID", _customerId);
                                cmd.Parameters.AddWithValue("@invID", invoiceId);
                                cmd.Parameters.AddWithValue("@amt", DebtPayment);
                                cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("o"));
                                cmd.ExecuteNonQuery();
                            }
                        }

                        // 6. Log AuditTrail
                        string insertAudit = @"
                            INSERT INTO AuditTrail (ActionType, EntityType, EntityID, Description, OldValue, NewValue, UserName, Timestamp)
                            VALUES ('INVOICE_LOCKED', 'Invoice', @invID, 'Invoice finalized and locked', '0', '1', 'admin', @now);";
                        using (var cmd = new SqliteCommand(insertAudit, conn, trans))
                        {
                            cmd.Parameters.AddWithValue("@invID", invoiceId);
                            cmd.Parameters.AddWithValue("@now", DateTime.UtcNow.ToString("o"));
                            cmd.ExecuteNonQuery();
                        }
                    });

                    // Lock UI
                    IsLocked = true;
                    LockStatusText = "🔒 الفاتورة مقفلة / Invoice Locked";

                    // 7. Dynamic QR Code generation
                    string qrPayload = $"INV:{InvoiceNo}|AMT:{GrandTotal}|CUST:{CustomerName}";

                    // 8. Silent Printing
                    PrintEngine.PrintInvoiceSilently(this, qrPayload);

                    // 9. PDF Archival
                    string pdfPath = PdfArchiver.ArchiveInvoicePdf(this, qrPayload);

                    // 10. Write-back stock quantities to stock2026.xlsx
                    if (!string.IsNullOrEmpty(_stockFilePath) && System.IO.File.Exists(_stockFilePath))
                    {
                        try
                        {
                            var deductions = InvoiceLines
                                .Where(l => !string.IsNullOrWhiteSpace(l.SourceServiceKey) && l.Qty > 0)
                                .GroupBy(l => l.SourceServiceKey!.Trim(), StringComparer.OrdinalIgnoreCase)
                                .ToDictionary(
                                    group => group.Key,
                                    group => group.Sum(line => line.Qty),
                                    StringComparer.OrdinalIgnoreCase);

                            if (deductions.Count > 0)
                                QYachtMaster.Inventory.PartsSyncEngine.WriteBackStockDeductions(_stockFilePath, deductions);
                        }
                        catch (Exception exWb)
                        {
                            // Non-fatal: log but don't block the user
                            MessageBox.Show($"تحذير: فشل تحديث ملف المخزون: {exWb.Message}\nWarning: Could not update stock file: {exWb.Message}",
                                "تحذير / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }

                    MessageBox.Show(
                        $"✅ تم إصدار الفاتورة وطباعتها وحفظها بنجاح\n" +
                        $"PDF: {pdfPath}\n\nInvoice issued, printed & archived.",
                        "نجاح / Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Navigate back to Main screen (close this window, open fresh MainWindow)
                    if (parameter is Window w)
                    {
                        var mainWindow = new MainWindow();
                        mainWindow.Show();
                        w.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"فشل قفل الفاتورة أو الطباعة / Save invoice failed: {ex.Message}", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    public class InvoiceLineViewModel : ViewModelBase
    {
        private readonly InvoiceViewModel _parent;
        private int _index;
        private string _description = string.Empty;
        private double _qty = 1.0;
        private double _unitPrice = 0.0;
        private double _discount = 0.0;
        private string? _discountNotes;
        private double _finalLineTotal = 0.0;
        private bool _isManualOverride = false;

        public int Index
        {
            get => _index;
            set => SetProperty(ref _index, value);
        }

        public string Description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

        public double Qty
        {
            get => _qty;
            set
            {
                if (SetProperty(ref _qty, value))
                {
                    Recalculate();
                }
            }
        }

        public double UnitPrice
        {
            get => _unitPrice;
            set
            {
                if (SetProperty(ref _unitPrice, value))
                {
                    Recalculate();
                }
            }
        }

        public double QtySubTotal => Qty * UnitPrice;

        public double Discount
        {
            get => _discount;
            set
            {
                if (SetProperty(ref _discount, value))
                {
                    Recalculate();
                }
            }
        }

        public string? DiscountNotes
        {
            get => _discountNotes;
            set => SetProperty(ref _discountNotes, value);
        }

        public double FinalLineTotal
        {
            get => _finalLineTotal;
            set
            {
                if (SetProperty(ref _finalLineTotal, value))
                {
                    IsManualOverride = true;
                    _parent.RecalculateTotals();
                }
            }
        }

        public bool IsManualOverride
        {
            get => _isManualOverride;
            set => SetProperty(ref _isManualOverride, value);
        }

        public double UnitCostPrice { get; set; } = 0.0;
        public string? SourceServiceKey { get; set; }

        private string? _warranty;
        public string? Warranty
        {
            get => _warranty;
            set => SetProperty(ref _warranty, value);
        }

        public InvoiceLineViewModel(InvoiceViewModel parent)
        {
            _parent = parent;
        }

        private void Recalculate()
        {
            _finalLineTotal = Math.Max(0.0, QtySubTotal - Discount);
            IsManualOverride = false;
            OnPropertyChanged(nameof(FinalLineTotal));
            OnPropertyChanged(nameof(QtySubTotal));
            _parent.RecalculateTotals();
        }
    }
}
