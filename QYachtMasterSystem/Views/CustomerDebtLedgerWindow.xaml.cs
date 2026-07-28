using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Data.Sqlite;
using QYachtMaster.Database;

namespace QYachtMaster.Views
{
    public partial class CustomerDebtLedgerWindow : Window
    {
        public class LedgerItem
        {
            public string InvoiceNo { get; set; } = string.Empty;
            public string Date { get; set; } = string.Empty;
            public double GrandTotal { get; set; }
            public double PaidAmount { get; set; }
            public double RemainingDebt { get; set; }
            public string Status { get; set; } = string.Empty;
        }

        public CustomerDebtLedgerWindow(int customerId, string customerName, string customerPhone, double totalDebt)
        {
            InitializeComponent();
            TxtCustomerName.Text = customerName;
            TxtCustomerPhone.Text = $"Mobile: {customerPhone}";
            TxtTotalDebt.Text = $"{totalDebt:N2} QAR";

            LoadLedgerData(customerId);
        }

        private void LoadLedgerData(int customerId)
        {
            var items = new List<LedgerItem>();
            string query = @"
                SELECT InvoiceNo, CreatedAt, GrandTotal, PaidAmount, RemainingDebt 
                FROM Invoices 
                WHERE CustomerID = @cID 
                ORDER BY InvoiceID DESC;";

            var param = new SqliteParameter("@cID", customerId);
            items = DatabaseHelper.ExecuteQuery(query, reader => new LedgerItem
            {
                InvoiceNo = reader["InvoiceNo"].ToString()!,
                Date = reader["CreatedAt"].ToString()!,
                GrandTotal = Convert.ToDouble(reader["GrandTotal"]),
                PaidAmount = Convert.ToDouble(reader["PaidAmount"]),
                RemainingDebt = Convert.ToDouble(reader["RemainingDebt"]),
                Status = Convert.ToDouble(reader["RemainingDebt"]) > 0 ? "⚠️ Debt Pending / دين معلق" : "✔️ Fully Paid / مدفوع بالكامل"
            }, param);

            GridLedger.ItemsSource = items;
        }

        private void OnCloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
