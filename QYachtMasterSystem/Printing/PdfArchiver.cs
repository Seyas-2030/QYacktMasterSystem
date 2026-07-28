using QYachtMaster.ViewModels;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace QYachtMaster.Printing
{
    public static class PdfArchiver
    {
        public static string ArchiveInvoicePdf(InvoiceViewModel viewModel, string qrPayload)
        {
            // ✨ التعديل الجديد: تفعيل مرشد الخطوط هنا في أول الدالة قبل استخدام أي خطوط
            try
            {
                if (PdfSharp.Fonts.GlobalFontSettings.FontResolver == null)
                {
                    PdfSharp.Fonts.GlobalFontSettings.FontResolver = new InvoiceFontResolver();
                }
            }
            catch { /* تم التعيين مسبقاً */ }

            PdfDocument document = new PdfDocument();
            document.Info.Title = $"Q-Yacht Master Invoice {viewModel.InvoiceNo}";

            PdfPage page = document.AddPage();
            page.Size = PdfSharp.PageSize.A4;
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Fonts
            XFont titleFont = new XFont("Arial", 18);
            XFont subtitleFont = new XFont("Arial", 10);
            XFont sectionFont = new XFont("Arial", 12);
            XFont regularFont = new XFont("Arial", 10);
            XFont boldFont = new XFont("Arial", 10);

            // Draw Header
            gfx.DrawString("Q-YACHT MASTER ENTERPRISE", titleFont, XBrushes.DarkRed, new XRect(40, 40, page.Width.Point - 80, 30), XStringFormats.TopLeft);
            gfx.DrawString("Lusail Marina, Doha, Qatar | C.R.: 123123 | Tel: +974 7226 1467", subtitleFont, XBrushes.DarkGray, new XRect(40, 70, page.Width.Point - 80, 20), XStringFormats.TopLeft);
            gfx.DrawString("INVOICE / فاتورة", titleFont, XBrushes.Black, new XRect(40, 95, page.Width.Point - 80, 30), XStringFormats.TopRight);

            // Divider Line
            gfx.DrawLine(XPens.LightGray, 40, 130, page.Width.Point - 40, 130);

            // Invoice details
            double yOffset = 145;
            gfx.DrawString($"Invoice No: {viewModel.InvoiceNo}", boldFont, XBrushes.Black, 40, yOffset);
            gfx.DrawString($"Date: {viewModel.InvoiceDate}", regularFont, XBrushes.Black, 350, yOffset);

            yOffset += 20;
            gfx.DrawString($"Customer Name: {viewModel.CustomerName}", regularFont, XBrushes.Black, 40, yOffset);
            gfx.DrawString($"Mobile: {viewModel.MobilePhone}", regularFont, XBrushes.Black, 350, yOffset);

            yOffset += 20;
            gfx.DrawString($"Vehicle: {viewModel.VehicleDetails}", regularFont, XBrushes.Black, 40, yOffset);

            // Grid header
            yOffset += 35;
            double gridTop = yOffset;
            gfx.DrawRectangle(XBrushes.LightSlateGray, 40, gridTop, page.Width.Point - 80, 22);

            gfx.DrawString("#", boldFont, XBrushes.White, 45, gridTop + 5);
            gfx.DrawString("Description", boldFont, XBrushes.White, 80, gridTop + 5);
            gfx.DrawString("Qty", boldFont, XBrushes.White, 350, gridTop + 5);
            gfx.DrawString("Unit Price", boldFont, XBrushes.White, 410, gridTop + 5);
            gfx.DrawString("Total (QAR)", boldFont, XBrushes.White, 500, gridTop + 5);

            yOffset += 22;

            // Lines
            foreach (var line in viewModel.InvoiceLines)
            {
                gfx.DrawString(line.Index.ToString(), regularFont, XBrushes.Black, 45, yOffset + 5);

                // Clean Arabic/English representation in descriptions for drawing
                string cleanDesc = line.Description;
                gfx.DrawString(cleanDesc, regularFont, XBrushes.Black, 80, yOffset + 5);

                gfx.DrawString(line.Qty.ToString("F1"), regularFont, XBrushes.Black, 350, yOffset + 5);
                gfx.DrawString(line.UnitPrice.ToString("F2"), regularFont, XBrushes.Black, 410, yOffset + 5);
                gfx.DrawString(line.FinalLineTotal.ToString("F2"), regularFont, XBrushes.Black, 500, yOffset + 5);

                yOffset += 20;
                gfx.DrawLine(XPens.Lavender, 40, yOffset, page.Width.Point - 40, yOffset);
            }

            // Summary
            yOffset += 15;
            double summaryRight = page.Width.Point - 40;

            gfx.DrawString("Sub-Total:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.SubTotal:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Discounts:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.TotalDiscounts:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Labor Fees:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.LaborTotal:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Previous Debt:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.PreviousDebt:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Debt Payment:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.DebtPayment:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Grand Total:", titleFont, XBrushes.Navy, 350, yOffset);
            gfx.DrawString($"{viewModel.GrandTotal:F2} QAR", titleFont, XBrushes.Navy, 480, yOffset);

            yOffset += 25;
            gfx.DrawString("Amount Paid:", boldFont, XBrushes.Black, 350, yOffset);
            gfx.DrawString($"{viewModel.PaidAmount:F2} QAR", regularFont, XBrushes.Black, 480, yOffset);

            yOffset += 18;
            gfx.DrawString("Balance Due:", boldFont, XBrushes.Crimson, 350, yOffset);
            gfx.DrawString($"{viewModel.BalanceDue:F2} QAR", boldFont, XBrushes.Crimson, 480, yOffset);

            // Draw QR Code
            try
            {
                byte[] qrPng = PrintEngine.GenerateQrCodePng(qrPayload);
                using (var ms = new MemoryStream(qrPng))
                {
                    XImage qrImage = XImage.FromStream(ms);
                    gfx.DrawImage(qrImage, 40, yOffset - 80, 100, 100);
                }
            }
            catch { /* Ignore QR drawing failures if stream error */ }

            // Save PDF
            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string year = DateTime.Now.ToString("yyyy");
            string month = DateTime.Now.ToString("MM");
            
            string rootFolder = Path.Combine(docPath, "QYachtMaster");
            string invoicesFolder = Path.Combine(rootFolder, "Invoices");
            string folder = Path.Combine(invoicesFolder, year, month);

            // Ensure rootFolder (QYachtMaster) exists and is visible (not hidden/system)
            if (!Directory.Exists(rootFolder))
            {
                Directory.CreateDirectory(rootFolder);
            }
            else
            {
                var di = new DirectoryInfo(rootFolder);
                di.Attributes &= ~FileAttributes.Hidden;
                di.Attributes &= ~FileAttributes.System;
            }

            // Create invoicesFolder (Invoices) and make it Hidden and System
            if (!Directory.Exists(invoicesFolder))
            {
                var di = Directory.CreateDirectory(invoicesFolder);
                di.Attributes |= FileAttributes.Hidden | FileAttributes.System;
            }
            else
            {
                var di = new DirectoryInfo(invoicesFolder);
                di.Attributes |= FileAttributes.Hidden | FileAttributes.System;
            }

            // Create year/month subfolder (visible/normal)
            if (!Directory.Exists(folder))
            {
                var di = Directory.CreateDirectory(folder);
                di.Attributes &= ~FileAttributes.Hidden;
                di.Attributes &= ~FileAttributes.System;
            }
            else
            {
                var di = new DirectoryInfo(folder);
                di.Attributes &= ~FileAttributes.Hidden;
                di.Attributes &= ~FileAttributes.System;
            }

            // Sanitize customer name for path
            string cleanCust = Regex.Replace(viewModel.CustomerName, @"[^a-zA-Z0-9\s]", "");
            cleanCust = cleanCust.Replace(" ", "_");

            string fileName = $"INV-{viewModel.InvoiceNo}-{cleanCust}.pdf";
            string fullPath = Path.Combine(folder, fileName);

            // 🔒 Full Encryption for PDF Invoice (requires password "0000" to open or run)
            document.SecuritySettings.UserPassword = "0000";
            document.SecuritySettings.OwnerPassword = "0000";
            document.SecuritySettings.PermitPrint = true;
            document.SecuritySettings.PermitModifyDocument = false;

            document.Save(fullPath);
            return fullPath;
        }
    }

    // class مساعد لتوجيه مكتبة PDFsharp إلى خطوط نظام التشغيل ويندوز
    public class InvoiceFontResolver : PdfSharp.Fonts.IFontResolver
    {
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName.Equals("Arial", StringComparison.OrdinalIgnoreCase))
            {
                if (isBold && isItalic) return new FontResolverInfo("Arial#bolditalic");
                if (isBold) return new FontResolverInfo("Arial#bold");
                if (isItalic) return new FontResolverInfo("Arial#italic");
                return new FontResolverInfo("Arial#regular");
            }
            return new FontResolverInfo("Arial#regular"); // خط افتراضي في حال اختيار خط آخر
        }

        public byte[] GetFont(string faceName)
        {
            // تحديد مجلد الخطوط الرسمي في نظام الويندوز تلقائياً
            string fontsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");

            switch (faceName.ToLower())
            {
                case "arial#regular":
                    return File.ReadAllBytes(Path.Combine(fontsFolder, "arial.ttf"));
                case "arial#bold":
                    return File.ReadAllBytes(Path.Combine(fontsFolder, "arialbd.ttf"));
                case "arial#italic":
                    return File.ReadAllBytes(Path.Combine(fontsFolder, "ariali.ttf"));
                case "arial#bolditalic":
                    return File.ReadAllBytes(Path.Combine(fontsFolder, "arialbi.ttf"));
                default:
                    return File.ReadAllBytes(Path.Combine(fontsFolder, "arial.ttf"));
            }
        }
    }
}