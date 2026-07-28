using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using QYachtMaster.Inventory;
using QYachtMaster.Database;

namespace QYachtMaster.Views
{
    public partial class ExcelSyncWindow : Window
    {
        private readonly List<PartsSyncEngine.ParsedPartInfo> _existingParts;
        private readonly List<PartsSyncEngine.ParsedPartInfo> _newParts;

        public ExcelSyncWindow(List<PartsSyncEngine.ParsedPartInfo> existingParts, List<PartsSyncEngine.ParsedPartInfo> newParts)
        {
            InitializeComponent();
            
            _existingParts = existingParts;
            _newParts = newParts;

            NewPartsGrid.ItemsSource = _newParts;
            SummaryText.Text = $"وجدنا {existingParts.Count} قطعة موجودة و {newParts.Count} قطعة جديدة / Found {existingParts.Count} existing parts and {newParts.Count} new parts.";
        }

        private void OnStartSyncClick(object sender, RoutedEventArgs e)
        {
            var selectedNewParts = _newParts.Where(p => p.IsSelectedToImport).ToList();

            try
            {
                // Execute bulk import/update database sync
                PartsSyncEngine.CommitSyncToDatabase(_existingParts, selectedNewParts);

                // Log audit trail event
                string desc = $"Excel synchronization: Updated {_existingParts.Count} existing parts, Imported {selectedNewParts.Count} new parts";
                string auditQuery = @"
                    INSERT INTO AuditTrail (ActionType, EntityType, Description, OldValue, NewValue, UserName, Timestamp)
                    VALUES ('INVENTORY_SYNC', 'Inventory', @desc, @oldVal, @newVal, 'admin', @time);";

                var ap = new[]
                {
                    new Microsoft.Data.Sqlite.SqliteParameter("@desc", desc),
                    new Microsoft.Data.Sqlite.SqliteParameter("@oldVal", $"Existing count: {_existingParts.Count}"),
                    new Microsoft.Data.Sqlite.SqliteParameter("@newVal", $"Imported count: {selectedNewParts.Count}"),
                    new Microsoft.Data.Sqlite.SqliteParameter("@time", DateTime.UtcNow.ToString("o"))
                };
                DatabaseHelper.ExecuteNonQuery(auditQuery, ap);

                MessageBox.Show($"تمت مزامنة المخزن بنجاح!\nتحديث: {_existingParts.Count} قطعة.\nاستيراد: {selectedNewParts.Count} قطعة جديدة.\nInventory sync finished successfully.", "نجاح / Success", MessageBoxButton.OK, MessageBoxImage.Information);
                
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل مزامنة المخزن / Database commit failed: {ex.Message}", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
