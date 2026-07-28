using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using QRCoder;

namespace QYachtMaster.Printing
{
    public class ThermalReceiptData
    {
        public string HeaderTitle { get; set; } = "⚓ Q-Yacht Master Enterprise";
        public string InvoiceNo { get; set; } = string.Empty;
        public string CustomerName { get; set; } = string.Empty;
        public string MobilePhone { get; set; } = string.Empty;
        public string VesselRegNo { get; set; } = string.Empty;
        public string DateText { get; set; } = DateTime.Now.ToString("g");
        public double SubTotal { get; set; }
        public double Discount { get; set; }
        public double GrandTotal { get; set; }
        public double PaidAmount { get; set; }
        public double RemainingDebt { get; set; }
        public string CashierName { get; set; } = "Admin";
    }

    public static class ThermalPrintEngine
    {
        public static void PrintThermalReceipt(ThermalReceiptData data, string printerName = "")
        {
            try
            {
                var pd = new PrintDocument();
                if (!string.IsNullOrEmpty(printerName))
                {
                    pd.PrinterSettings.PrinterName = printerName;
                }

                pd.PrintPage += (sender, e) =>
                {
                    Graphics g = e.Graphics!;
                    Font headerFont = new Font("Courier New", 12, FontStyle.Bold);
                    Font titleFont = new Font("Courier New", 10, FontStyle.Bold);
                    Font bodyFont = new Font("Courier New", 8, FontStyle.Regular);
                    Font boldBodyFont = new Font("Courier New", 8, FontStyle.Bold);

                    float startX = 5;
                    float startY = 10;
                    float offsetY = 0;

                    g.DrawString(data.HeaderTitle, headerFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 20;
                    g.DrawString("------------------------------", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 15;

                    g.DrawString($"Invoice: {data.InvoiceNo}", titleFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 16;
                    g.DrawString($"Date: {data.DateText}", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Customer: {data.CustomerName}", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Phone: {data.MobilePhone}", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Vessel Reg: {data.VesselRegNo}", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 18;

                    g.DrawString("------------------------------", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 15;

                    g.DrawString($"SubTotal:      {data.SubTotal:N2} QAR", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Discount:      {data.Discount:N2} QAR", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Grand Total:   {data.GrandTotal:N2} QAR", boldBodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 16;
                    g.DrawString($"Paid Amount:   {data.PaidAmount:N2} QAR", bodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 14;
                    g.DrawString($"Remaining:     {data.RemainingDebt:N2} QAR", boldBodyFont, Brushes.Black, startX, startY + offsetY);
                    offsetY += 20;

                    // Generate QR Code for Receipt
                    using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                    using (QRCodeData qrCodeData = qrGenerator.CreateQrCode($"QYACHT|{data.InvoiceNo}|{data.GrandTotal}", QRCodeGenerator.ECCLevel.Q))
                    using (BitmapByteQRCode qrCode = new BitmapByteQRCode(qrCodeData))
                    {
                        byte[] qrBytes = qrCode.GetGraphic(3);
                        using (MemoryStream ms = new MemoryStream(qrBytes))
                        using (Bitmap qrBitmap = new Bitmap(ms))
                        {
                            g.DrawImage(qrBitmap, startX + 50, startY + offsetY, 80, 80);
                        }
                    }
                    offsetY += 90;

                    g.DrawString("Thank you for choosing Q-Yacht!", bodyFont, Brushes.Black, startX + 10, startY + offsetY);
                };

                pd.Print();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ThermalPrintEngine Exception: {ex.Message}");
            }
        }
    }
}
