using System.Windows;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class PartReturnWindow : Window
    {
        private readonly PartReturnViewModel _vm;

        /// <summary>
        /// Opens the Part Return dialog for a specific customer.
        /// </summary>
        /// <param name="customerID">The CustomerID whose invoice is being reversed.</param>
        /// <param name="customerName">Display name shown in the header.</param>
        /// <param name="currentUsername">The logged-in employee processing the return.</param>
        public PartReturnWindow(int customerID, string customerName, string currentUsername)
        {
            // Create the VM before initializing components so any events raised
            // during InitializeComponent (e.g. Checked handlers) see a valid _vm.
            _vm = new PartReturnViewModel(customerID, customerName, currentUsername);

            InitializeComponent();

            DataContext = _vm;

            // Close this dialog automatically when the VM fires the success event
            _vm.ReturnCompleted += (_, _) => Close();
        }

        // ── Radio button handlers → update VM's RefundType string ────────
        private void CashRadio_Checked(object sender, RoutedEventArgs e)
            => _vm.RefundType = "Cash";

        private void DebtRadio_Checked(object sender, RoutedEventArgs e)
            => _vm.RefundType = "DebtReduction";

        private void OnCancelClick(object sender, RoutedEventArgs e)
            => Close();
    }
}
