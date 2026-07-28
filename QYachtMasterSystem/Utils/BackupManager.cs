using System;
using System.IO;
using Microsoft.Data.Sqlite;
using QYachtMaster.Database;

namespace QYachtMaster.Utils
{
    /// <summary>
    /// Automated local backup engine with VACUUM INTO and cold-storage archiving.
    /// Backs up the live database to NAS / secondary drive and archives old records.
    /// </summary>
    public static class BackupManager
    {
        /// <summary>
        /// Performs a live, zero-downtime backup using VACUUM INTO.
        /// Outputs a compressed, defragmented copy of the main database.
        /// </summary>
        public static string BackupToPath(string? destinationFolder = null)
        {
            if (string.IsNullOrEmpty(destinationFolder))
            {
                string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                destinationFolder = Path.Combine(docPath, "QYachtMaster", "Backups");
            }

            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupFile = Path.Combine(destinationFolder, $"qyachtmaster_backup_{timestamp}.db");

            string connStr = DatabaseInitializer.GetConnectionString();
            using (var connection = new SqliteConnection(connStr))
            {
                connection.Open();
                using (var cmd = new SqliteCommand($"VACUUM INTO @dest;", connection))
                {
                    cmd.Parameters.AddWithValue("@dest", backupFile);
                    cmd.ExecuteNonQuery();
                }
            }

            return backupFile;
        }

        /// <summary>
        /// Archives closed invoices older than the specified number of years
        /// into a separate cold-storage database file to keep operational tables lean.
        /// </summary>
        public static int ArchiveOldInvoices(int yearsOld = 3, string? archiveFolder = null)
        {
            if (string.IsNullOrEmpty(archiveFolder))
            {
                string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                archiveFolder = Path.Combine(docPath, "QYachtMaster", "Archive");
            }

            if (!Directory.Exists(archiveFolder))
            {
                Directory.CreateDirectory(archiveFolder);
            }

            string archiveDb = Path.Combine(archiveFolder, "qyachtmaster_archive.db");
            string cutoffDate = DateTime.UtcNow.AddYears(-yearsOld).ToString("o");

            int archivedCount = 0;

            string mainConnStr = DatabaseInitializer.GetConnectionString();
            string archiveConnStr = $"Data Source={archiveDb}";

            // Create archive DB schema if needed
            using (var archiveConn = new SqliteConnection(archiveConnStr))
            {
                archiveConn.Open();
                using (var cmd = new SqliteCommand(@"
                    CREATE TABLE IF NOT EXISTS ArchivedInvoices (
                        InvoiceID     INTEGER PRIMARY KEY,
                        InvoiceNo     TEXT    NOT NULL,
                        JobCardID     INTEGER NOT NULL,
                        CustomerID    INTEGER NOT NULL,
                        SubTotal      REAL    NOT NULL DEFAULT 0.0,
                        LaborTotal    REAL    NOT NULL DEFAULT 0.0,
                        GrandTotal    REAL    NOT NULL DEFAULT 0.0,
                        PaidAmount    REAL    NOT NULL DEFAULT 0.0,
                        RemainingDebt REAL    NOT NULL DEFAULT 0.0,
                        PaymentMethod TEXT,
                        CreatedAt     TEXT    NOT NULL,
                        ArchivedAt    TEXT    NOT NULL
                    );", archiveConn))
                {
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = new SqliteCommand(@"
                    CREATE TABLE IF NOT EXISTS ArchivedInvoiceLines (
                        LineID            INTEGER PRIMARY KEY,
                        InvoiceID         INTEGER NOT NULL,
                        Description       TEXT    NOT NULL,
                        Qty               REAL    NOT NULL DEFAULT 1.0,
                        UnitPrice         REAL    NOT NULL DEFAULT 0.0,
                        FinalLineTotal    REAL    NOT NULL DEFAULT 0.0,
                        ArchivedAt        TEXT    NOT NULL
                    );", archiveConn))
                {
                    cmd.ExecuteNonQuery();
                }
            }

            // Copy old locked invoices to archive then delete from main
            using (var mainConn = new SqliteConnection(mainConnStr))
            {
                mainConn.Open();

                // Enable WAL on main connection
                using (var cmd = new SqliteCommand("PRAGMA journal_mode = WAL;", mainConn))
                    cmd.ExecuteNonQuery();

                using (var archiveConn = new SqliteConnection(archiveConnStr))
                {
                    archiveConn.Open();

                    // Get old invoices
                    string selectSql = @"
                        SELECT InvoiceID, InvoiceNo, JobCardID, CustomerID,
                               SubTotal, LaborTotal, GrandTotal, PaidAmount, RemainingDebt,
                               PaymentMethod, CreatedAt
                        FROM Invoices
                        WHERE IsLocked = 1 AND CreatedAt < @cutoff;";

                    using (var selectCmd = new SqliteCommand(selectSql, mainConn))
                    {
                        selectCmd.Parameters.AddWithValue("@cutoff", cutoffDate);
                        using (var reader = selectCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int invoiceId = reader.GetInt32(0);
                                string now = DateTime.UtcNow.ToString("o");

                                // Insert into archive
                                using (var insertCmd = new SqliteCommand(@"
                                    INSERT OR IGNORE INTO ArchivedInvoices
                                    (InvoiceID, InvoiceNo, JobCardID, CustomerID, SubTotal, LaborTotal,
                                     GrandTotal, PaidAmount, RemainingDebt, PaymentMethod, CreatedAt, ArchivedAt)
                                    VALUES (@id, @no, @jid, @cid, @sub, @lab, @grand, @paid, @rem, @pay, @created, @archived);",
                                    archiveConn))
                                {
                                    insertCmd.Parameters.AddWithValue("@id", invoiceId);
                                    insertCmd.Parameters.AddWithValue("@no", reader.GetString(1));
                                    insertCmd.Parameters.AddWithValue("@jid", reader.GetInt32(2));
                                    insertCmd.Parameters.AddWithValue("@cid", reader.GetInt32(3));
                                    insertCmd.Parameters.AddWithValue("@sub", reader.GetDouble(4));
                                    insertCmd.Parameters.AddWithValue("@lab", reader.GetDouble(5));
                                    insertCmd.Parameters.AddWithValue("@grand", reader.GetDouble(6));
                                    insertCmd.Parameters.AddWithValue("@paid", reader.GetDouble(7));
                                    insertCmd.Parameters.AddWithValue("@rem", reader.GetDouble(8));
                                    insertCmd.Parameters.AddWithValue("@pay", reader[9]?.ToString() ?? "");
                                    insertCmd.Parameters.AddWithValue("@created", reader.GetString(10));
                                    insertCmd.Parameters.AddWithValue("@archived", now);
                                    insertCmd.ExecuteNonQuery();
                                }

                                archivedCount++;
                            }
                        }
                    }

                    // Delete archived invoices and their lines from main DB
                    if (archivedCount > 0)
                    {
                        using (var trans = mainConn.BeginTransaction())
                        {
                            try
                            {
                                using (var delLines = new SqliteCommand(
                                    "DELETE FROM InvoiceLines WHERE InvoiceID IN (SELECT InvoiceID FROM Invoices WHERE IsLocked = 1 AND CreatedAt < @cutoff);",
                                    mainConn, trans))
                                {
                                    delLines.Parameters.AddWithValue("@cutoff", cutoffDate);
                                    delLines.ExecuteNonQuery();
                                }

                                using (var delInvoices = new SqliteCommand(
                                    "DELETE FROM Invoices WHERE IsLocked = 1 AND CreatedAt < @cutoff;",
                                    mainConn, trans))
                                {
                                    delInvoices.Parameters.AddWithValue("@cutoff", cutoffDate);
                                    delInvoices.ExecuteNonQuery();
                                }

                                trans.Commit();
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }
                }
            }

            return archivedCount;
        }

        /// <summary>
        /// Cleans up old backup files, keeping only the most recent N backups.
        /// </summary>
        public static void PruneOldBackups(int keepCount = 10, string? backupFolder = null)
        {
            if (string.IsNullOrEmpty(backupFolder))
            {
                string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                backupFolder = Path.Combine(docPath, "QYachtMaster", "Backups");
            }

            if (!Directory.Exists(backupFolder)) return;

            var files = new DirectoryInfo(backupFolder)
                .GetFiles("qyachtmaster_backup_*.db")
                .OrderByDescending(f => f.CreationTime)
                .ToArray();

            for (int i = keepCount; i < files.Length; i++)
            {
                try { files[i].Delete(); } catch { }
            }
        }
    }
}
