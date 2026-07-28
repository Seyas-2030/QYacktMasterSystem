using System;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using QYachtMaster.ViewModels;

namespace QYachtMaster.Views
{
    public partial class PrintInvoiceLayout : UserControl
    {
        public PrintInvoiceLayout(InvoiceViewModel viewModel, BitmapImage qrImage)
        {
            InitializeComponent();

            // Populate metadata
            InvoiceDateText.Text = viewModel.InvoiceDate;
            InvoiceNoText.Text = $"No. {viewModel.InvoiceNo}";
            CustomerNameText.Text = $"Mr. / Messrs. / السيد السادة:  {viewModel.CustomerName}";
            CustomerPhoneText.Text = $"Phone / Mobile / الجوال:  {viewModel.MobilePhone}";

            // Vehicle Details
            string[] vehParts = viewModel.VehicleDetails.Split('-');
            VehRegText.Text = $"REG. NO / رقم اللوحة:  {(vehParts.Length > 0 ? vehParts[0].Trim() : string.Empty)}";
            VehMakeText.Text = $"MAKE / الماركة:  {(vehParts.Length > 1 ? vehParts[1].Trim() : string.Empty)}";
            VehModelText.Text = $"MODEL / الموديل:  {(vehParts.Length > 2 ? vehParts[2].Trim() : string.Empty)}";
            VehYearText.Text = "YEAR / سنة الصنع:  ";

            // Financial Summary
            SubTotalText.Text = $"{viewModel.SubTotal:F2} QAR";
            DiscountText.Text = $"{viewModel.TotalDiscounts:F2} QAR";
            LaborText.Text = $"{viewModel.LaborTotal:F2} QAR";
            PreviousDebtText.Text = $"{viewModel.PreviousDebt:F2} QAR";
            DebtPaymentText.Text = $"{viewModel.DebtPayment:F2} QAR";
            GrandTotalText.Text = $"{viewModel.GrandTotal:F2} QAR";
            PaidAmountText.Text = $"{viewModel.PaidAmount:F2} QAR";
            BalanceDueText.Text = $"{viewModel.BalanceDue:F2} QAR";

            // Set QR code source
            QrCodeImage.Source = qrImage;

            // Bind lines items list
            ItemsList.ItemsSource = viewModel.InvoiceLines;
        }
    }
}
