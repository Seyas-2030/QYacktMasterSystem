using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Xunit;
using QYachtMaster.Database;
using QYachtMaster.Inventory;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Tests
{
    public class BusinessLogicTests
    {
        public BusinessLogicTests()
        {
            // Initialize database schema so that tables like JobCards exist during unit test execution
            DatabaseInitializer.Initialize();
        }

        [Theory]
        [InlineData("1pc", "1")]
        [InlineData("2PC", "2")]
        [InlineData("12.5", "12.5")]
        [InlineData("-3", "-3")]
        [InlineData("abc", "0")]
        [InlineData("", "0")]
        [InlineData(null, "0")]
        public void CleanQtyText_ShouldSanitizeValuesCorrectly(string? input, string expected)
        {
            // Act
            string actual = PartsSyncEngine.CleanQtyText(input ?? string.Empty);

            // Assert
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void InvoiceLineViewModel_ShouldCalculateTotalsCorrectlyWithoutOverride()
        {
            // Arrange
            var parentViewModel = new InvoiceViewModel(0); // Pass a dummy jobCardId
            var line = new InvoiceLineViewModel(parentViewModel)
            {
                Qty = 2.0,
                UnitPrice = 50.0,
                Discount = 15.0
            };

            // Assert
            Assert.Equal(100.0, line.QtySubTotal);
            Assert.Equal(85.0, line.FinalLineTotal);
            Assert.False(line.IsManualOverride);
        }

        [Fact]
        public void InvoiceLineViewModel_ShouldDetectManualOverride()
        {
            // Arrange
            var parentViewModel = new InvoiceViewModel(0);
            var line = new InvoiceLineViewModel(parentViewModel)
            {
                Qty = 3.0,
                UnitPrice = 100.0,
                Discount = 20.0
            };

            // Act - Manually override the final total
            line.FinalLineTotal = 250.0;

            // Assert
            Assert.True(line.IsManualOverride);
            Assert.Equal(250.0, line.FinalLineTotal);
            Assert.Equal(300.0, line.QtySubTotal);
        }

        [Fact]
        public void WriteBackStockDeductions_ShouldUpdateSharedWorkbook()
        {
            string filePath = Path.Combine(Path.GetTempPath(), $"qym-stock-{Guid.NewGuid():N}.xlsx");

            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("QYachtMaster");
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Stock");
                    sheet.Cells[1, 1].Value = "PART NO";
                    sheet.Cells[1, 2].Value = "QTY";
                    sheet.Cells[2, 1].Value = "PART-001";
                    sheet.Cells[2, 2].Value = 10;
                    package.SaveAs(new FileInfo(filePath));
                }

                // Simulate another application, such as Excel, keeping a shared handle open.
                using (var sharedHandle = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    PartsSyncEngine.WriteBackStockDeductions(filePath, new Dictionary<string, double>
                    {
                        ["PART-001"] = 3
                    });
                }

                using var updatedPackage = new ExcelPackage(new FileInfo(filePath));
                Assert.Equal(7d, updatedPackage.Workbook.Worksheets["Stock"].Cells[2, 2].GetValue<double>());
            }
            finally
            {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void DatabaseInitializer_ShouldSetDatabasePathAndInitializeSchema()
        {
            string dbPath = DatabaseInitializer.GetDatabasePath();
            Assert.Contains("QYachtMaster", dbPath);
            Assert.True(File.Exists(dbPath));

            string journalMode = DatabaseHelper.ExecuteScalar("PRAGMA journal_mode;").ToString()!;
            Assert.Equal("wal", journalMode.ToLower());
        }

        [Fact]
        public void BackupManager_ShouldCreateVacuumBackupSuccessfully()
        {
            string backupPath = QYachtMaster.Utils.BackupManager.BackupToPath();
            Assert.True(File.Exists(backupPath));
            Assert.True(new FileInfo(backupPath).Length > 0);

            // Clean test file
            try { File.Delete(backupPath); } catch { }
        }

        [Fact]
        public void LanSyncEngine_ShouldFormatMessageCorrectly()
        {
            var msg = new QYachtMaster.Utils.LanSyncMessage
            {
                EventType = "BAY_CHANGED",
                EntityID = 101,
                Payload = "Bay 2: Electrical"
            };

            Assert.Equal("BAY_CHANGED", msg.EventType);
            Assert.Equal(101, msg.EntityID);
            Assert.Equal("Bay 2: Electrical", msg.Payload);
            Assert.Equal(Environment.MachineName, msg.SenderMachine);
        }

        [Fact]
        public async System.Threading.Tasks.Task StreamImportPartsAsync_ShouldStreamImportPartsToDatabase()
        {
            var parts = new List<PartsSyncEngine.ParsedPartInfo>
            {
                new PartsSyncEngine.ParsedPartInfo
                {
                    PartNo = $"TEST-PART-{Guid.NewGuid():N}",
                    Description = "Marine Pump Impeller",
                    Category = "Engine",
                    Prefix = "EP",
                    QtyInExcel = 5,
                    SealValue = 0,
                    SalePrice = 120.0,
                    CostPrice = 75.0,
                    IsSelectedToImport = true
                }
            };

            await PartsSyncEngine.StreamImportPartsAsync(parts);

            var dbPart = DatabaseHelper.ExecuteSingleQuery(
                "SELECT * FROM Parts WHERE PartNo = @p;",
                reader => reader["Description"]?.ToString(),
                new Microsoft.Data.Sqlite.SqliteParameter("@p", parts[0].PartNo));

            Assert.Equal("Marine Pump Impeller", dbPart);
        }
    }
}
