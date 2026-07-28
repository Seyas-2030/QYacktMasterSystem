using System.Windows;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class InvoiceWindow : Window
    {
        public InvoiceWindow(int jobCardId, string? stockFilePath = null)
        {
            InitializeComponent();
            DataContext = new InvoiceViewModel(jobCardId, stockFilePath);
        }
    }
}
