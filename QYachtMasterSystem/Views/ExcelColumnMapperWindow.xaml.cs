using System;
using System.Data;
using System.Windows;

namespace QYachtMaster.Views
{
    public partial class ExcelColumnMapperWindow : Window
    {
        public string PartNoColumn { get; private set; } = string.Empty;
        public string QtyColumn { get; private set; } = string.Empty;
        public string SalePriceColumn { get; private set; } = string.Empty;
        public string DescriptionColumn { get; private set; } = string.Empty;

        public ExcelColumnMapperWindow(DataTable dataTable)
        {
            InitializeComponent();

            foreach (DataColumn column in dataTable.Columns)
            {
                PartNoCombo.Items.Add(column.ColumnName);
                QtyCombo.Items.Add(column.ColumnName);
                SalePriceCombo.Items.Add(column.ColumnName);
                DescriptionCombo.Items.Add(column.ColumnName);
            }

            // Guess mappings based on standard headers
            GuessMappings();
        }

        private void GuessMappings()
        {
            // Auto-select matchings
            SelectItemContains(PartNoCombo, "part");
            SelectItemContains(QtyCombo, "qty");
            SelectItemContains(QtyCombo, "quantity");
            SelectItemContains(SalePriceCombo, "price");
            SelectItemContains(SalePriceCombo, "sale");
            SelectItemContains(DescriptionCombo, "desc");
        }

        private void SelectItemContains(System.Windows.Controls.ComboBox combo, string keyword)
        {
            foreach (var item in combo.Items)
            {
                if (item.ToString()!.ToLower().Contains(keyword.ToLower()))
                {
                    combo.SelectedItem = item;
                    break;
                }
            }
        }

        private void OnImportClick(object sender, RoutedEventArgs e)
        {
            if (PartNoCombo.SelectedItem == null || QtyCombo.SelectedItem == null)
            {
                MessageBox.Show("يرجى اختيار أعمدة رقم القطعة والكمية الإجبارية / Please map both Part Number and Quantity columns.", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            PartNoColumn = PartNoCombo.SelectedItem.ToString()!;
            QtyColumn = QtyCombo.SelectedItem.ToString()!;
            SalePriceColumn = SalePriceCombo.SelectedItem?.ToString() ?? string.Empty;
            DescriptionColumn = DescriptionCombo.SelectedItem?.ToString() ?? string.Empty;

            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
