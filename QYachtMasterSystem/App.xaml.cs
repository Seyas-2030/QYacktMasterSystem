using System;
using System.Threading.Tasks;
using System.Windows;
using QYachtMaster.Database;
using QYachtMaster.Reports;
using QYachtMaster.Utils;

namespace QYachtMaster
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Initialize local SQLite database
            DatabaseInitializer.Initialize();

            // Run auto-backup and auto-monthly-report in background without blocking UI
            Task.Run(() =>
            {
                try
                {
                    // 1. Auto Backup on startup
                    string backupFile = BackupManager.BackupToPath();
                    BackupManager.PruneOldBackups(keepCount: 10);
                }
                catch { /* Silent fail — don't block app startup */ }

                try
                {
                    // 2. Auto Monthly Report — runs if a new month has started since last export
                    var (nextYear, nextMonth) = ReportExporter.DetermineNextMonthlyPeriod();
                    var now = DateTime.Now;
                    // Only auto-export if the determined period is before or equal to last month
                    // (meaning we haven't exported the previous month's report yet)
                    bool shouldExport = (nextYear < now.Year) ||
                                        (nextYear == now.Year && nextMonth < now.Month);
                    if (shouldExport)
                    {
                        ReportExporter.ExportMonthlyReportAuto(out _, out _);
                    }
                }
                catch { /* Silent fail */ }
            });
        }
    }
}
