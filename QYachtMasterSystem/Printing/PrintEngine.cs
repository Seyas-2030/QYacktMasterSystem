using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QRCoder;
using QYachtMaster.ViewModels;
using QYachtMaster.Views;

namespace QYachtMaster.Printing
{
    public static class PrintEngine
    {
        public static byte[] GenerateQrCodePng(string text)
        {
            using (var qrGenerator = new QRCodeGenerator())
            using (var qrCodeData = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q))
            using (var qrCode = new PngByteQRCode(qrCodeData))
            {
                return qrCode.GetGraphic(5);
            }
        }

        public static BitmapImage ByteArrayToImage(byte[] bytes)
        {
            var bitmap = new BitmapImage();
            using (var stream = new MemoryStream(bytes))
            {
                bitmap.BeginInit();
                bitmap.StreamSource = stream;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
            }
            return bitmap;
        }

        public static void PrintInvoiceSilently(InvoiceViewModel viewModel, string qrPayload)
        {
            try
            {
                // Generate QR code Image
                byte[] qrBytes = GenerateQrCodePng(qrPayload);
                BitmapImage qrImage = ByteArrayToImage(qrBytes);

                // Create visual layout
                var printLayout = new PrintInvoiceLayout(viewModel, qrImage);

                // Set size to standard A4 (96 DPI: 794 x 1123 pixels)
                printLayout.Width = 794;
                printLayout.Height = 1123;

                // Measure and arrange layout
                printLayout.Measure(new Size(794, 1123));
                printLayout.Arrange(new Rect(0, 0, 794, 1123));
                printLayout.UpdateLayout();

                // Spool silently to system default printer
                PrintDialog dialog = new PrintDialog();
                dialog.PrintVisual(printLayout, $"Q-Yacht Master Invoice {viewModel.InvoiceNo}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"فشل الطباعة الصامتة / Silent printing failed: {ex.Message}", "خطأ / Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
