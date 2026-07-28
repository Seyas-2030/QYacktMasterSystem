using System;
using System.Windows;
using System.Windows.Input;
using Microsoft.Data.Sqlite;
using QYachtMaster.MVVM;
using QYachtMaster.Database;
using QYachtMaster.Models;

namespace QYachtMaster.ViewModels
{
    public class PartReturnViewModel : ViewModelBase
    {
        // ── Injected context ─────────────────────────────────────────────
        private readonly int    _customerID;
        private readonly string _customerName;
        public  readonly string CurrentUsername;

        // ── Fields ───────────────────────────────────────────────────────
        private string _invoiceNo    = string.Empty;
        private string _partNo       = string.Empty;
        private double _qtyReturned  = 1.0;
        private double _unitPrice    = 0.0;
        private double _refundAmount = 0.0;
        private string _refundType   = "Cash";       // Cash | DebtReduction
        private string _notes        = string.Empty;
        private string _statusMsg    = string.Empty;
        private string _statusColor  = "Transparent";
        private bool   _partFound    = false;
        private int?   _resolvedPartID;

        // ── Properties ───────────────────────────────────────────────────
        public string CustomerDisplay => $"{_customerName}  ({_customerID})";

        public string InvoiceNo
        {
            get => _invoiceNo;
            set => SetProperty(ref _invoiceNo, value);
        }

        public string PartNo
        {
            get => _partNo;
            set
            {
                if (SetProperty(ref _partNo, value))
                    LookupPart();
            }
        }

        public double QtyReturned
        {
            get => _qtyReturned;
            set
            {
                if (SetProperty(ref _qtyReturned, value))
                    RecalcRefund();
            }
        }

        public double UnitPrice
        {
            get => _unitPrice;
            set
            {
                if (SetProperty(ref _unitPrice, value))
                    RecalcRefund();
            }
        }

        public double RefundAmount
        {
            get => _refundAmount;
            set => SetProperty(ref _refundAmount, value);
        }

        public string RefundType
        {
            get => _refundType;
            set => SetProperty(ref _refundType, value);
        }

        public string Notes
        {
            get => _notes;
            set => SetProperty(ref _notes, value);
        }

        public string StatusMsg
        {
            get => _statusMsg;
            set => SetProperty(ref _statusMsg, value);
        }

        public string StatusColor
        {
            get => _statusColor;
            set => SetProperty(ref _statusColor, value);
        }

        public bool PartFound
        {
            get => _partFound;
            set => SetProperty(ref _partFound, value);
        }

        // ── Commands ─────────────────────────────────────────────────────
        public ICommand LookupPartCommand   { get; }
        public ICommand ConfirmReturnCommand { get; }

        // ── Constructor ──────────────────────────────────────────────────
        public PartReturnViewModel(int customerID, string customerName, string currentUsername)
        {
            _customerID   = customerID;
            _customerName = customerName;
            CurrentUsername = currentUsername;

            LookupPartCommand    = new RelayCommand(_ => LookupPart());
            ConfirmReturnCommand = new RelayCommand(_ => ExecuteReturn(), _ => CanConfirm());
        }

        // ── Part lookup ──────────────────────────────────────────────────
        private void LookupPart()
        {
            if (string.IsNullOrWhiteSpace(PartNo))
            {
                PartFound    = false;
                _resolvedPartID = null;
                StatusMsg    = string.Empty;
                StatusColor  = "Transparent";
                return;
            }

            var part = DatabaseHelper.ExecuteSingleQuery(
                "SELECT PartID, SalePrice FROM Parts WHERE LOWER(PartNo) = LOWER(@pno);",
                r => new { PartID = Convert.ToInt32(r["PartID"]), SalePrice = Convert.ToDouble(r["SalePrice"]) },
                new SqliteParameter("@pno", PartNo.Trim()));

            if (part != null)
            {
                _resolvedPartID = part.PartID;
                UnitPrice       = part.SalePrice;
                RecalcRefund();
                PartFound    = true;
                StatusMsg    = $"✅ Part found — Sale price: {part.SalePrice:F2} QAR";
                StatusColor  = "#27AE60";
            }
            else
            {
                _resolvedPartID = null;
                PartFound    = false;
                StatusMsg    = "❌ Part number not found in database / رقم القطعة غير موجود";
                StatusColor  = "#E74C3C";
            }
        }

        private void RecalcRefund() =>
            RefundAmount = Math.Round(QtyReturned * UnitPrice, 2);

        private bool CanConfirm() =>
            PartFound && QtyReturned > 0 && RefundAmount >= 0;

        // ── Core transaction ─────────────────────────────────────────────
        private void ExecuteReturn()
        {
            if (!CanConfirm()) return;

            string partNoCleaned = PartNo.Trim();
            string invoiceNoCleaned = InvoiceNo.Trim();
            string notesCleaned = Notes.Trim();
            string now = DateTime.UtcNow.ToString("o");

            // Build automatic notes if employee left field blank
            string autoNote = string.IsNullOrEmpty(notesCleaned)
                ? $"إرجاع {QtyReturned} × {partNoCleaned} — فاتورة: {invoiceNoCleaned}"
                : notesCleaned;

            try
            {
                DatabaseHelper.ExecuteTransaction((conn, trans) =>
                {
                    // ── 1. Restore part qty to stock ──────────────────────────
                    using (var cmd = new SqliteCommand(
                        "UPDATE Parts SET QtyInStock = QtyInStock + @qty WHERE PartNo = @pno;",
                        conn, trans))
                    {
                        cmd.Parameters.AddWithValue("@qty", QtyReturned);
                        cmd.Parameters.AddWithValue("@pno", partNoCleaned);
                        cmd.ExecuteNonQuery();
                    }

                    // ── 2a. Cash Refund → negative Payment ───────────────────
                    if (RefundType == "Cash")
                    {
                        using var cmd = new SqliteCommand(@"
                            INSERT INTO Payments
                                (CustomerID, InvoiceID, Amount, PaymentType, ReceivedBy, ReceivedAt, Notes)
                            VALUES
                                (@cid, @invID, @amt, 'Cash', @user, @at, @notes);",
                            conn, trans);

                        // Resolve InvoiceID if the employee provided an invoice number
                        object? invIdObj = ResolveInvoiceID(invoiceNoCleaned, conn, trans);

                        cmd.Parameters.AddWithValue("@cid",   _customerID);
                        cmd.Parameters.AddWithValue("@invID", invIdObj ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@amt",   -RefundAmount);   // NEGATIVE = cash out
                        cmd.Parameters.AddWithValue("@user",  CurrentUsername);
                        cmd.Parameters.AddWithValue("@at",    now);
                        cmd.Parameters.AddWithValue("@notes", $"PART_RETURN | {autoNote}");
                        cmd.ExecuteNonQuery();
                    }
                    // ── 2b. Debt Reduction → lower TotalDebt ─────────────────
                    else
                    {
                        using var cmd = new SqliteCommand(
                            "UPDATE Customers SET TotalDebt = MAX(0, TotalDebt - @amt) WHERE CustomerID = @cid;",
                            conn, trans);
                        cmd.Parameters.AddWithValue("@amt", RefundAmount);
                        cmd.Parameters.AddWithValue("@cid", _customerID);
                        cmd.ExecuteNonQuery();
                    }

                    // ── 3. Save PartReturns record ────────────────────────────
                    using (var cmd = new SqliteCommand(@"
                        INSERT INTO PartReturns
                            (InvoiceID, InvoiceNo, CustomerID, PartID, PartNo,
                             QtyReturned, UnitPrice, RefundAmount, RefundType,
                             ProcessedBy, ProcessedAt, Notes)
                        VALUES
                            (@invID, @invNo, @cid, @partID, @partNo,
                             @qty, @price, @refund, @rtype,
                             @user, @at, @notes);",
                        conn, trans))
                    {
                        object? invIdObj2 = ResolveInvoiceID(invoiceNoCleaned, conn, trans);
                        cmd.Parameters.AddWithValue("@invID",  invIdObj2 ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@invNo",  invoiceNoCleaned);
                        cmd.Parameters.AddWithValue("@cid",    _customerID);
                        cmd.Parameters.AddWithValue("@partID", _resolvedPartID.HasValue ? (object)_resolvedPartID.Value : DBNull.Value);
                        cmd.Parameters.AddWithValue("@partNo", partNoCleaned);
                        cmd.Parameters.AddWithValue("@qty",    QtyReturned);
                        cmd.Parameters.AddWithValue("@price",  UnitPrice);
                        cmd.Parameters.AddWithValue("@refund", RefundAmount);
                        cmd.Parameters.AddWithValue("@rtype",  RefundType);
                        cmd.Parameters.AddWithValue("@user",   CurrentUsername);
                        cmd.Parameters.AddWithValue("@at",     now);
                        cmd.Parameters.AddWithValue("@notes",  autoNote);
                        cmd.ExecuteNonQuery();
                    }

                    // ── 4. AuditTrail ─────────────────────────────────────────
                    using (var cmd = new SqliteCommand(@"
                        INSERT INTO AuditTrail
                            (ActionType, EntityType, EntityID, Description,
                             OldValue, NewValue, UserName, Timestamp)
                        VALUES
                            ('PART_RETURN', 'Customer', @cid, @desc,
                             @old, @new, @user, @time);",
                        conn, trans))
                    {
                        cmd.Parameters.AddWithValue("@cid",  _customerID);
                        cmd.Parameters.AddWithValue("@desc",
                            $"Part returned: {QtyReturned}× {partNoCleaned} | Refund: {RefundAmount:F2} QAR ({RefundType}) | Inv: {invoiceNoCleaned}");
                        cmd.Parameters.AddWithValue("@old",  $"QtyInStock before +{QtyReturned}");
                        cmd.Parameters.AddWithValue("@new",  $"RefundAmount={RefundAmount:F2}, RefundType={RefundType}");
                        cmd.Parameters.AddWithValue("@user", CurrentUsername);
                        cmd.Parameters.AddWithValue("@time", now);
                        cmd.ExecuteNonQuery();
                    }
                });

                // ── 5. Write-back stock returns to Excel ──────────────────────
                string? stockFilePath = DatabaseInitializer.GetSetting("StockFilePath");
                if (!string.IsNullOrEmpty(stockFilePath) && System.IO.File.Exists(stockFilePath))
                {
                    try
                    {
                        var returns = new System.Collections.Generic.Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
                        {
                            { partNoCleaned, QtyReturned }
                        };
                        QYachtMaster.Inventory.PartsSyncEngine.WriteBackStockReturns(stockFilePath, returns);
                    }
                    catch (Exception exWb)
                    {
                        // Non-fatal warning
                        MessageBox.Show($"تحذير: فشل تحديث ملف المخزون Excel: {exWb.Message}\nWarning: Could not update Excel stock file: {exWb.Message}",
                            "تحذير / Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }


                // ── Success ───────────────────────────────────────────────
                string successAr = RefundType == "Cash"
                    ? $"تم خصم {RefundAmount:F2} ر.ق من الصندوق وإرجاع الكمية للمخزن."
                    : $"تم خصم {RefundAmount:F2} ر.ق من مديونية العميل وإرجاع الكمية للمخزن.";

                MessageBox.Show(
                    $"✅ Return processed successfully!\n{successAr}\n\n" +
                    $"  • Part: {partNoCleaned}\n" +
                    $"  • Qty Returned: {QtyReturned}\n" +
                    $"  • Refund: {RefundAmount:F2} QAR  [{RefundType}]",
                    "Return Successful / تم الاسترجاع",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Close the dialog via event
                ReturnCompleted?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"فشل الاسترجاع / Return failed:\n{ex.Message}",
                    "Error / خطأ",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>Tries to resolve an InvoiceNo string to a numeric InvoiceID. Returns null if not found.</summary>
        private static object? ResolveInvoiceID(string invoiceNo, SqliteConnection conn, SqliteTransaction trans)
        {
            if (string.IsNullOrWhiteSpace(invoiceNo)) return null;
            using var cmd = new SqliteCommand(
                "SELECT InvoiceID FROM Invoices WHERE InvoiceNo = @no LIMIT 1;", conn, trans);
            cmd.Parameters.AddWithValue("@no", invoiceNo);
            var result = cmd.ExecuteScalar();
            return result is long id ? (object)id : null;
        }

        /// <summary>Raised when the return transaction completes successfully. The Window can subscribe to close itself.</summary>
        public event EventHandler? ReturnCompleted;
    }
}
