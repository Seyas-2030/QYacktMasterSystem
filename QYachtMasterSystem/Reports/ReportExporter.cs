using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using QYachtMaster.Database;
using Microsoft.Data.Sqlite;

namespace QYachtMaster.Reports
{
    public static class ReportExporter
    {
        // ─── Data Models ───────────────────────────────────────────────
        public class PartsProfitRow
        {
            public string Description { get; set; } = string.Empty;
            public string SourceServiceKey { get; set; } = string.Empty;
            public double TotalQtySold { get; set; }
            public double TotalRevenue { get; set; }
            public double TotalCost { get; set; }
            public double NetProfit { get; set; }
        }

        public class CustomerDebtRow
        {
            public string Name { get; set; } = string.Empty;
            public string MobilePhone { get; set; } = string.Empty;
            public double TotalDebt { get; set; }
            public string LastJobDate { get; set; } = string.Empty;
        }

        public class TopSellingPartRow
        {
            public string Description { get; set; } = string.Empty;
            public double TotalQtySold { get; set; }
            public double TotalRevenue { get; set; }
        }

        public class LowStockPartRow
        {
            public string PartNo { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public double QtyInStock { get; set; }
            public string Location { get; set; } = string.Empty;
        }

        // ─── Data Fetchers ──────────────────────────────────────────────
        public static List<PartsProfitRow> FetchPartsProfitability()
        {
            // Line Net Profit = FinalLineTotal - (UnitCostPrice × Qty)
            string query = @"
                SELECT
                    il.Description,
                    il.SourceServiceKey,
                    SUM(il.Qty)                                 AS TotalQtySold,
                    SUM(il.FinalLineTotal)                      AS TotalRevenue,
                    SUM(il.UnitCostPrice * il.Qty)              AS TotalCost,
                    SUM(il.FinalLineTotal - il.UnitCostPrice * il.Qty) AS NetProfit
                FROM InvoiceLines il
                INNER JOIN Invoices i ON il.InvoiceID = i.InvoiceID
                WHERE i.IsLocked = 1
                GROUP BY il.Description
                ORDER BY NetProfit DESC;";

            return DatabaseHelper.ExecuteQuery(query, reader => new PartsProfitRow
            {
                Description      = reader["Description"]?.ToString() ?? string.Empty,
                SourceServiceKey = reader["SourceServiceKey"]?.ToString() ?? string.Empty,
                TotalQtySold     = Convert.ToDouble(reader["TotalQtySold"]),
                TotalRevenue     = Convert.ToDouble(reader["TotalRevenue"]),
                TotalCost        = Convert.ToDouble(reader["TotalCost"]),
                NetProfit        = Convert.ToDouble(reader["NetProfit"])
            });
        }

        public static List<CustomerDebtRow> FetchCustomerDebts()
        {
            string query = @"
                SELECT c.Name, c.MobilePhone, c.TotalDebt,
                       MAX(j.JobDate) AS LastJobDate
                FROM Customers c
                LEFT JOIN JobCards j ON j.CustomerID = c.CustomerID
                WHERE c.TotalDebt > 0
                GROUP BY c.CustomerID
                ORDER BY c.TotalDebt DESC;";

            return DatabaseHelper.ExecuteQuery(query, reader => new CustomerDebtRow
            {
                Name         = reader["Name"]?.ToString() ?? string.Empty,
                MobilePhone  = reader["MobilePhone"]?.ToString() ?? string.Empty,
                TotalDebt    = Convert.ToDouble(reader["TotalDebt"]),
                LastJobDate  = reader["LastJobDate"]?.ToString() ?? string.Empty
            });
        }

        public static List<TopSellingPartRow> FetchTopSellingParts(int topN = 20)
        {
            string query = $@"
                SELECT il.Description,
                       SUM(il.Qty) AS TotalQtySold,
                       SUM(il.FinalLineTotal) AS TotalRevenue
                FROM InvoiceLines il
                INNER JOIN Invoices i ON il.InvoiceID = i.InvoiceID
                WHERE i.IsLocked = 1
                GROUP BY il.Description
                ORDER BY TotalQtySold DESC
                LIMIT {topN};";

            return DatabaseHelper.ExecuteQuery(query, reader => new TopSellingPartRow
            {
                Description  = reader["Description"]?.ToString() ?? string.Empty,
                TotalQtySold = Convert.ToDouble(reader["TotalQtySold"]),
                TotalRevenue = Convert.ToDouble(reader["TotalRevenue"])
            });
        }

        public static List<LowStockPartRow> FetchLowStockParts(double threshold = 3.0)
        {
            string query = @"
                SELECT PartNo, Description, Category, QtyInStock, Location
                FROM Parts
                WHERE QtyInStock <= @threshold
                ORDER BY QtyInStock ASC;";

            var p = new SqliteParameter("@threshold", threshold);
            return DatabaseHelper.ExecuteQuery(query, reader => new LowStockPartRow
            {
                PartNo      = reader["PartNo"]?.ToString() ?? string.Empty,
                Description = reader["Description"]?.ToString() ?? string.Empty,
                Category    = reader["Category"]?.ToString() ?? string.Empty,
                QtyInStock  = Convert.ToDouble(reader["QtyInStock"]),
                Location    = reader["Location"]?.ToString() ?? string.Empty
            }, p);
        }

        // ─── Excel Exporters ────────────────────────────────────────────
        public static string ExportPartsProfitability()
        {
            var data = FetchPartsProfitability();
            double totalNetProfit = data.Sum(r => r.NetProfit);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Parts Profitability");

            var headers = new[] { "Description / البيان", "Service Key", "Qty Sold",
                                   "Revenue (QAR)", "Cost (QAR)", "Net Profit (QAR)" };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var r in data)
            {
                ws.Cell(row, 1).Value = r.Description;
                ws.Cell(row, 2).Value = r.SourceServiceKey;
                ws.Cell(row, 3).Value = r.TotalQtySold;
                ws.Cell(row, 4).Value = r.TotalRevenue;
                ws.Cell(row, 5).Value = r.TotalCost;
                ws.Cell(row, 6).Value = r.NetProfit;
                var profitCell = ws.Cell(row, 6);
                profitCell.Style.Font.FontColor = r.NetProfit >= 0
                    ? XLColor.DarkGreen : XLColor.DarkRed;
                row++;
            }

            // Add the overall parts profit separately from each part's individual profit.
            var totalLabel = ws.Range(row, 1, row, 5);
            totalLabel.Merge();
            totalLabel.Value = "Total Parts Net Profit / إجمالي أرباح جميع القطع";
            totalLabel.Style.Font.Bold = true;
            totalLabel.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAD3");
            totalLabel.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var totalProfitCell = ws.Cell(row, 6);
            totalProfitCell.Value = totalNetProfit;
            totalProfitCell.Style.Font.Bold = true;
            totalProfitCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAD3");
            totalProfitCell.Style.Font.FontColor = totalNetProfit >= 0
                ? XLColor.DarkGreen : XLColor.DarkRed;

            ws.Columns().AdjustToContents();
            return SaveWorkbook(wb, $"PartsProfitability_{DateTime.Now:yyyyMMdd}");
        }

        public static string ExportCustomerDebts()
        {
            var data = FetchCustomerDebts();
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Customer Debts");

            var headers = new[] { "Customer Name / الاسم", "Mobile / الجوال",
                                   "Debt (QAR) / المديونية", "Last Visit / آخر زيارة" };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var r in data)
            {
                ws.Cell(row, 1).Value = r.Name;
                ws.Cell(row, 2).Value = r.MobilePhone;
                ws.Cell(row, 3).Value = r.TotalDebt;
                ws.Cell(row, 3).Style.Font.FontColor = XLColor.DarkRed;
                ws.Cell(row, 4).Value = r.LastJobDate;
                row++;
            }

            ws.Columns().AdjustToContents();
            return SaveWorkbook(wb, $"CustomerDebts_{DateTime.Now:yyyyMMdd}");
        }

        public static string ExportTopSellingParts()
        {
            var data = FetchTopSellingParts();
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Top Selling Parts");

            var headers = new[] { "Description / البيان", "Qty Sold / المبيعات",
                                   "Revenue (QAR) / الإيرادات" };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var r in data)
            {
                ws.Cell(row, 1).Value = r.Description;
                ws.Cell(row, 2).Value = r.TotalQtySold;
                ws.Cell(row, 3).Value = r.TotalRevenue;
                row++;
            }

            ws.Columns().AdjustToContents();
            return SaveWorkbook(wb, $"TopSellingParts_{DateTime.Now:yyyyMMdd}");
        }

        public static string ExportLowStockParts()
        {
            var data = FetchLowStockParts();
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Low Stock Alert");

            var headers = new[] { "Part No / رقم القطعة", "Description / الوصف",
                                   "Category / الصنف", "Stock / الرصيد", "Location / الموقع" };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var r in data)
            {
                ws.Cell(row, 1).Value = r.PartNo;
                ws.Cell(row, 2).Value = r.Description;
                ws.Cell(row, 3).Value = r.Category;
                ws.Cell(row, 4).Value = r.QtyInStock;
                var stockCell = ws.Cell(row, 4);
                stockCell.Style.Font.FontColor = r.QtyInStock == 0 ? XLColor.Red : XLColor.OrangeRed;
                ws.Cell(row, 5).Value = r.Location;
                row++;
            }

            ws.Columns().AdjustToContents();
            return SaveWorkbook(wb, $"LowStockAlert_{DateTime.Now:yyyyMMdd}");
        }


        // ─── Detailed Reports (No GROUP BY) ────────────────────────────
        public class DetailedReportRow
        {
            public string InvoiceNo { get; set; } = string.Empty;
            public string CreatedAt { get; set; } = string.Empty;
            public string CustomerName { get; set; } = string.Empty;
            public double SubTotal { get; set; }
            public double Discount { get; set; }
            public double GrandTotal { get; set; }
            public double PartsCost { get; set; }
            public double NetProfit { get; set; }
        }

        public static List<string> FetchAvailablePeriods(string periodType)
        {
            // Only monthly is supported as yearly/daily/weekly are cancelled/removed.
            string expr = "strftime('%Y-%m', CreatedAt)";

            string query = $@"
                SELECT DISTINCT {expr} AS Period 
                FROM Invoices 
                WHERE IsLocked = 1 
                ORDER BY Period DESC;";

            return DatabaseHelper.ExecuteQuery(query, reader => reader["Period"]?.ToString() ?? string.Empty)
                                 .Where(s => !string.IsNullOrEmpty(s))
                                 .ToList();
        }

        public static List<DetailedReportRow> FetchDetailedReport(string periodType, string periodValue)
        {
            // Only monthly is supported now
            string filterExpr = "strftime('%Y-%m', i.CreatedAt) = @val";

            string query = $@"
                SELECT 
                    i.InvoiceNo, 
                    i.CreatedAt, 
                    c.Name AS CustomerName, 
                    i.SubTotal, 
                    (i.SubTotal + i.LaborTotal - i.GrandTotal) AS Discount, 
                    i.GrandTotal,
                    IFNULL(cost.TotalCost, 0) AS PartsCost,
                    (i.GrandTotal - IFNULL(cost.TotalCost, 0)) AS NetProfit
                FROM Invoices i
                LEFT JOIN Customers c ON i.CustomerID = c.CustomerID
                LEFT JOIN (
                    SELECT InvoiceID, SUM(UnitCostPrice * Qty) AS TotalCost
                    FROM InvoiceLines
                    GROUP BY InvoiceID
                ) cost ON i.InvoiceID = cost.InvoiceID
                WHERE i.IsLocked = 1
                  AND {filterExpr}
                ORDER BY i.InvoiceNo ASC;";

            var p = new SqliteParameter("@val", periodValue);
            return DatabaseHelper.ExecuteQuery(query, reader => new DetailedReportRow
            {
                InvoiceNo    = reader["InvoiceNo"]?.ToString() ?? string.Empty,
                CreatedAt    = reader["CreatedAt"]?.ToString() ?? string.Empty,
                CustomerName = reader["CustomerName"]?.ToString() ?? string.Empty,
                SubTotal     = Convert.ToDouble(reader["SubTotal"]),
                Discount     = Convert.ToDouble(reader["Discount"]),
                GrandTotal   = Convert.ToDouble(reader["GrandTotal"]),
                PartsCost    = Convert.ToDouble(reader["PartsCost"]),
                NetProfit    = Convert.ToDouble(reader["NetProfit"])
            }, p);
        }

        public static string ExportMonthlyReportAuto(out int exportedYear, out int exportedMonth)
        {
            var (year, month) = DetermineNextMonthlyPeriod();
            exportedYear = year;
            exportedMonth = month;

            string periodValue = $"{year}-{month:D2}";
            var data = FetchDetailedReport("monthly", periodValue);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Detailed Report");

            var headers = new[] 
            { 
                "Invoice No / رقم الفاتورة", 
                "Date & Time / التاريخ والوقت", 
                "Customer Name / اسم العميل", 
                "Sub-Total / المجموع الفرعي", 
                "Discount / الخصم", 
                "Grand Total / الإجمالي الكلي", 
                "Parts Cost / تكلفة القطع", 
                "Invoice Profit / صافي الربح" 
            };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            double sumSubTotal = 0;
            double sumDiscount = 0;
            double sumGrandTotal = 0;
            double sumPartsCost = 0;
            double sumNetProfit = 0;

            if (data.Count == 0)
            {
                ws.Cell(row, 1).Value = "No invoices found for this month / لا توجد فواتير مسجلة لهذا الشهر";
                ws.Range(row, 1, row, 8).Merge();
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(row, 1).Style.Font.Italic = true;
                ws.Cell(row, 1).Style.Font.FontColor = XLColor.Gray;
                row++;
            }
            else
            {
                foreach (var r in data)
                {
                    ws.Cell(row, 1).Value = r.InvoiceNo;
                    ws.Cell(row, 2).Value = r.CreatedAt;
                    ws.Cell(row, 3).Value = r.CustomerName;
                    ws.Cell(row, 4).Value = r.SubTotal;
                    ws.Cell(row, 5).Value = r.Discount;
                    ws.Cell(row, 6).Value = r.GrandTotal;
                    ws.Cell(row, 7).Value = r.PartsCost;
                    ws.Cell(row, 8).Value = r.NetProfit;

                    var profitCell = ws.Cell(row, 8);
                    profitCell.Style.Font.FontColor = r.NetProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;

                    sumSubTotal   += r.SubTotal;
                    sumDiscount   += r.Discount;
                    sumGrandTotal += r.GrandTotal;
                    sumPartsCost  += r.PartsCost;
                    sumNetProfit  += r.NetProfit;
                    row++;
                }

                ws.Cell(row, 1).Value = "Total / الإجمالي";
                ws.Cell(row, 1).Style.Font.Bold = true;
                
                ws.Cell(row, 4).Value = sumSubTotal;
                ws.Cell(row, 4).Style.Font.Bold = true;

                ws.Cell(row, 5).Value = sumDiscount;
                ws.Cell(row, 5).Style.Font.Bold = true;

                ws.Cell(row, 6).Value = sumGrandTotal;
                ws.Cell(row, 6).Style.Font.Bold = true;

                ws.Cell(row, 7).Value = sumPartsCost;
                ws.Cell(row, 7).Style.Font.Bold = true;

                ws.Cell(row, 8).Value = sumNetProfit;
                ws.Cell(row, 8).Style.Font.Bold = true;
                ws.Cell(row, 8).Style.Font.FontColor = sumNetProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;

                var totalRange = ws.Range(row, 1, row, 8);
                totalRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                totalRange.Style.Border.BottomBorder = XLBorderStyleValues.Double;
                totalRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAEAEA");
            }

            // Month 12 / December Year-End Summary Highlight
            if (month == 12)
            {
                row += 2; // Leave a blank row
                
                // December Net Profit
                var decProfitCellLabel = ws.Cell(row, 1);
                decProfitCellLabel.Value = "December Net Profit / صافي ربح هذا الشهر (ديسمبر)";
                decProfitCellLabel.Style.Font.Bold = true;
                decProfitCellLabel.Style.Font.FontSize = 11;
                
                var decProfitCellValue = ws.Cell(row, 3);
                decProfitCellValue.Value = sumNetProfit;
                decProfitCellValue.Style.Font.Bold = true;
                decProfitCellValue.Style.Font.FontSize = 11;
                decProfitCellValue.Style.Font.FontColor = sumNetProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;
                
                ws.Range(row, 1, row, 2).Merge();

                row++;

                // Total Yearly Net Profit
                double yearlyProfit = FetchYearlyNetProfit(year);
                var yearProfitCellLabel = ws.Cell(row, 1);
                yearProfitCellLabel.Value = $"Total Yearly Net Profit for {year} / صافي ربح السنة كاملة ({year})";
                yearProfitCellLabel.Style.Font.Bold = true;
                yearProfitCellLabel.Style.Font.FontSize = 12;
                
                var yearProfitCellValue = ws.Cell(row, 3);
                yearProfitCellValue.Value = yearlyProfit;
                yearProfitCellValue.Style.Font.Bold = true;
                yearProfitCellValue.Style.Font.FontSize = 12;
                yearProfitCellValue.Style.Font.FontColor = yearlyProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;
                
                ws.Range(row, 1, row, 2).Merge();

                // Style Year-End Summary
                var summaryRange = ws.Range(row - 1, 1, row, 3);
                summaryRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                summaryRange.Style.Border.OutsideBorderColor = XLColor.FromHtml("#2A394A");
                summaryRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F7FA");
            }

            ws.Columns().AdjustToContents();

            string subFolder = year.ToString();
            string baseName = $"Detailed_Report_monthly_Month_{month:D2}";
            return SaveWorkbook(wb, baseName, subFolder);
        }

        public static List<DetailedReportRow> FetchCustomDetailedReport(string startStr, string endStr)
        {
            string query = @"
                SELECT 
                    i.InvoiceNo, 
                    i.CreatedAt, 
                    c.Name AS CustomerName, 
                    i.SubTotal, 
                    (i.SubTotal + i.LaborTotal - i.GrandTotal) AS Discount, 
                    i.GrandTotal,
                    IFNULL(cost.TotalCost, 0) AS PartsCost,
                    (i.GrandTotal - IFNULL(cost.TotalCost, 0)) AS NetProfit
                FROM Invoices i
                LEFT JOIN Customers c ON i.CustomerID = c.CustomerID
                LEFT JOIN (
                    SELECT InvoiceID, SUM(UnitCostPrice * Qty) AS TotalCost
                    FROM InvoiceLines
                    GROUP BY InvoiceID
                ) cost ON i.InvoiceID = cost.InvoiceID
                WHERE i.IsLocked = 1
                  AND date(i.CreatedAt) >= date(@start)
                  AND date(i.CreatedAt) <= date(@end)
                ORDER BY i.InvoiceNo ASC;";

            var pStart = new SqliteParameter("@start", startStr);
            var pEnd = new SqliteParameter("@end", endStr);
            return DatabaseHelper.ExecuteQuery(query, reader => new DetailedReportRow
            {
                InvoiceNo    = reader["InvoiceNo"]?.ToString() ?? string.Empty,
                CreatedAt    = reader["CreatedAt"]?.ToString() ?? string.Empty,
                CustomerName = reader["CustomerName"]?.ToString() ?? string.Empty,
                SubTotal     = Convert.ToDouble(reader["SubTotal"]),
                Discount     = Convert.ToDouble(reader["Discount"]),
                GrandTotal   = Convert.ToDouble(reader["GrandTotal"]),
                PartsCost    = Convert.ToDouble(reader["PartsCost"]),
                NetProfit    = Convert.ToDouble(reader["NetProfit"])
            }, pStart, pEnd);
        }

        public static string ExportCustomDetailedReport(string startStr, string endStr)
        {
            var data = FetchCustomDetailedReport(startStr, endStr);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Detailed Report");

            var headers = new[] 
            { 
                "Invoice No / رقم الفاتورة", 
                "Date & Time / التاريخ والوقت", 
                "Customer Name / اسم العميل", 
                "Sub-Total / المجموع الفرعي", 
                "Discount / الخصم", 
                "Grand Total / الإجمالي الكلي", 
                "Parts Cost / تكلفة القطع", 
                "Invoice Profit / صافي الربح" 
            };
            ApplyHeaderRow(ws, 1, headers);

            int row = 2;
            double sumSubTotal = 0;
            double sumDiscount = 0;
            double sumGrandTotal = 0;
            double sumPartsCost = 0;
            double sumNetProfit = 0;

            foreach (var r in data)
            {
                ws.Cell(row, 1).Value = r.InvoiceNo;
                ws.Cell(row, 2).Value = r.CreatedAt;
                ws.Cell(row, 3).Value = r.CustomerName;
                ws.Cell(row, 4).Value = r.SubTotal;
                ws.Cell(row, 5).Value = r.Discount;
                ws.Cell(row, 6).Value = r.GrandTotal;
                ws.Cell(row, 7).Value = r.PartsCost;
                ws.Cell(row, 8).Value = r.NetProfit;

                var profitCell = ws.Cell(row, 8);
                profitCell.Style.Font.FontColor = r.NetProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;

                sumSubTotal   += r.SubTotal;
                sumDiscount   += r.Discount;
                sumGrandTotal += r.GrandTotal;
                sumPartsCost  += r.PartsCost;
                sumNetProfit  += r.NetProfit;
                row++;
            }

            // Total Row
            ws.Cell(row, 1).Value = "Total / الإجمالي";
            ws.Cell(row, 1).Style.Font.Bold = true;
            
            ws.Cell(row, 4).Value = sumSubTotal;
            ws.Cell(row, 4).Style.Font.Bold = true;

            ws.Cell(row, 5).Value = sumDiscount;
            ws.Cell(row, 5).Style.Font.Bold = true;

            ws.Cell(row, 6).Value = sumGrandTotal;
            ws.Cell(row, 6).Style.Font.Bold = true;

            ws.Cell(row, 7).Value = sumPartsCost;
            ws.Cell(row, 7).Style.Font.Bold = true;

            ws.Cell(row, 8).Value = sumNetProfit;
            ws.Cell(row, 8).Style.Font.Bold = true;
            ws.Cell(row, 8).Style.Font.FontColor = sumNetProfit >= 0 ? XLColor.DarkGreen : XLColor.DarkRed;

            var totalRange = ws.Range(row, 1, row, 8);
            totalRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            totalRange.Style.Border.BottomBorder = XLBorderStyleValues.Double;
            totalRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAEAEA");

            ws.Columns().AdjustToContents();

            string startClean = startStr.Replace("-", "_");
            string endClean = endStr.Replace("-", "_");
            string baseName = $"Detailed_Report_Custom_{startClean}_to_{endClean}";
            return SaveWorkbook(wb, baseName, "Custom");
        }

        // ─── Helpers ────────────────────────────────────────────────────
        public static (int Year, int Month) GetEarliestLockedInvoicePeriod()
        {
            string query = "SELECT MIN(CreatedAt) FROM Invoices WHERE IsLocked = 1;";
            object? val = DatabaseHelper.ExecuteScalar(query);
            if (val != DBNull.Value && val != null)
            {
                string dateStr = val.ToString() ?? string.Empty;
                if (DateTime.TryParse(dateStr, out DateTime dt))
                {
                    return (dt.Year, dt.Month);
                }
            }
            return (DateTime.Now.Year, DateTime.Now.Month);
        }

        public static (int Year, int Month) DetermineNextMonthlyPeriod()
        {
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string reportsRoot = Path.Combine(docPath, "QYachtMaster", "Reports");

            // Find all 4-digit numeric subfolders (Years)
            List<int> yearDirs = new List<int>();
            if (Directory.Exists(reportsRoot))
            {
                yearDirs = Directory.GetDirectories(reportsRoot)
                    .Select(d => Path.GetFileName(d))
                    .Where(name => name.Length == 4 && int.TryParse(name, out _))
                    .Select(name => int.Parse(name))
                    .OrderBy(y => y)
                    .ToList();
            }

            if (yearDirs.Count == 0)
            {
                // No year folders exist yet. Determine the start period based on earliest locked invoice
                return GetEarliestLockedInvoicePeriod();
            }

            int maxYear = yearDirs.Last();
            string maxYearFolder = Path.Combine(reportsRoot, maxYear.ToString());

            // Search files named: Detailed_Report_monthly_Month_*.xlsx
            var files = Directory.GetFiles(maxYearFolder, "Detailed_Report_monthly_Month_*.xlsx")
                .Select(f => Path.GetFileNameWithoutExtension(f))
                .ToList();

            if (files.Count == 0)
            {
                // If folder exists but has no reports, default to Month 1 of that year
                return (maxYear, 1);
            }

            var months = new List<int>();
            foreach (var f in files)
            {
                string prefix = "Detailed_Report_monthly_Month_";
                if (f.StartsWith(prefix) && f.Length >= prefix.Length + 2)
                {
                    string mStr = f.Substring(prefix.Length, 2);
                    if (int.TryParse(mStr, out int m))
                    {
                        months.Add(m);
                    }
                }
            }

            if (months.Count == 0)
            {
                return (maxYear, 1);
            }

            int maxMonth = months.Max();
            if (maxMonth < 12)
            {
                return (maxYear, maxMonth + 1);
            }
            else
            {
                return (maxYear + 1, 1);
            }
        }

        public static double FetchYearlyNetProfit(int year)
        {
            string query = @"
                SELECT SUM(i.GrandTotal - IFNULL(cost.TotalCost, 0)) AS YearlyNetProfit
                FROM Invoices i
                LEFT JOIN (
                    SELECT InvoiceID, SUM(UnitCostPrice * Qty) AS TotalCost
                    FROM InvoiceLines
                    GROUP BY InvoiceID
                ) cost ON i.InvoiceID = cost.InvoiceID
                WHERE i.IsLocked = 1
                  AND strftime('%Y', i.CreatedAt) = @year;";

            var param = new SqliteParameter("@year", year.ToString());
            object? val = DatabaseHelper.ExecuteScalar(query, param);
            return val != DBNull.Value && val != null ? Convert.ToDouble(val) : 0.0;
        }

        private static void ApplyHeaderRow(IXLWorksheet ws, int row, string[] headers)
        {
            for (int col = 1; col <= headers.Length; col++)
            {
                var cell = ws.Cell(row, col);
                cell.Value = headers[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#2A394A");
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
        }

        private static string SaveWorkbook(XLWorkbook wb, string baseName, string subFolder = "")
        {
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folder  = Path.Combine(docPath, "QYachtMaster", "Reports", subFolder);
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);

            string path = Path.Combine(folder, $"{baseName}.xlsx");
            wb.SaveAs(path);
            return path;
        }
    }
}
