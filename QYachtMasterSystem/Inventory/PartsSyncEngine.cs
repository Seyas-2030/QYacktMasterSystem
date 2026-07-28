using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Channels;
using System.Threading.Tasks;
using OfficeOpenXml;
using QYachtMaster.Database;
using QYachtMaster.Models;

namespace QYachtMaster.Inventory
{
    /// <summary>
    /// Handles all interactions with stock Excel files (password-protected, opened via EPPlus).
    /// ClosedXML is reserved for admin report generation only.
    /// </summary>
    public class PartsSyncEngine
    {
        // EPPlus 8+ Modern License configuration
        static PartsSyncEngine()
        {
            ExcelPackage.License.SetNonCommercialPersonal("QYachtMaster");
        }

        private const string ExcelPassword = "0000";

        public class ParsedPartInfo
        {
            public string PartNo { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public string Category { get; set; } = string.Empty;
            public string Prefix { get; set; } = string.Empty;
            public double QtyInExcel { get; set; }
            public double SealValue { get; set; }
            public double CalculatedQty => QtyInExcel + SealValue;
            public string Location { get; set; } = string.Empty;
            public string CarModel { get; set; } = string.Empty;
            public double CostPrice { get; set; } = 0.0;
            public double SalePrice { get; set; } = 0.0;
            public bool IsNew { get; set; }
            public bool IsSelectedToImport { get; set; } = true;
        }

        public class StockPartEntry
        {
            public string PartNo    { get; set; } = string.Empty;
            public string Model     { get; set; } = string.Empty;
            public string BoxLabel  { get; set; } = string.Empty;  // e.g. "Box 1", "Master Box"
            public double Qty       { get; set; }
            public double SalePrice { get; set; }
            public double CostPrice { get; set; }
        }

        // Sheet name → (Category label, prefix)
        private static readonly Dictionary<string, (string Category, string Prefix)> WorksheetSchema
            = new(StringComparer.OrdinalIgnoreCase)
        {
            { "EVPARATER",        ("Evaporator",        "EV") },
            { "HEATER CORE",      ("Heater Core",       "HT") },
            { "DARIYER FILTER",   ("Dryer Filter",      "RD") },
            { "AC COMORESSURE",   ("Compressor",        "CO") },
            { "HOUSE",            ("Heater Housing",    "HA") },
            { "EXPANSHAN VALVEL", ("Expansion Valve",   "EX") },
            { "SWICH",            ("Pressure Switch",   "SW") },
            { "FILTER",           ("AC Filter",         "FI") },
            { "CONDANSOER",       ("Condenser",         "CN") },
            { "FAN",              ("Fan/Blower Motor",  "FA") }
        };

        // ----------------------------------------------------------------
        // Utility
        // ----------------------------------------------------------------
        public static string CleanQtyText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "0";
            var cleaned = Regex.Replace(raw, @"[^0-9\-\.]", "");
            return string.IsNullOrEmpty(cleaned) ? "0" : cleaned;
        }

        private static bool IsStockHeaderOrCategory(string value)
        {
            string upper = value.Trim().ToUpperInvariant();
            return upper == "PART" || upper.StartsWith("PART NO") ||
                   upper.StartsWith("PARTS NO") || upper.StartsWith("PART NUMBER") ||
                   upper == "QTY" || upper == "QTV" || upper == "TOTAL" ||
                   upper == "HEATER" || upper == "EVAPORATOR" ||
                   upper == "FAN" || upper == "SWICH" || upper == "FILTER" ||
                   upper.Contains("AC COMP") || upper.Contains("CONDANSOER");
        }

        // ----------------------------------------------------------------
        // 1. Get sheet names from the password-protected Excel file
        // ----------------------------------------------------------------
        public static List<string> GetSheetNames(string filePath)
        {
            var names = new List<string>();
            using var pkg = OpenProtected(filePath);
            foreach (var ws in pkg.Workbook.Worksheets)
            {
                names.Add(ws.Name);
            }
            return names;
        }
        // ----------------------------------------------------------------
        // 2. Get all parts for a specific sheet/department (Dropdowns)
        // ----------------------------------------------------------------
        /// <summary>
        /// Returns all stockable parts from a worksheet for use in Job Card dropdowns.
        /// Supports two layouts automatically:
        ///   (a) VERTICAL   â€“ one part per row (most sheets).
        ///   (b) MULTI-BOX  â€“ horizontal groups headed by "BOX-01", "BOX-02", â€¦, "MASTER BOX"
        ///                    (HOUSE sheet).  The right-hand "price columns" (duplicate box
        ///                    labels with only PARTS NO + COST columns) are extracted into a
        ///                    price map and merged back onto the data rows.
        /// Every returned entry carries a BoxLabel for display in the dropdown.
        /// </summary>
        public static List<StockPartEntry> GetPartsForSheet(string filePath, string sheetName)
        {
            var results = new List<StockPartEntry>();
            using var pkg = OpenProtected(filePath);

            var ws = pkg.Workbook.Worksheets[sheetName];
            if (ws == null) return results;

            int rows = ws.Dimension?.Rows ?? 0;
            int cols = ws.Dimension?.Columns ?? 0;
            if (rows < 1 || cols < 1) return results;

            // Explicitly force DARIYER FILTER and SWICH sheets to be parsed vertically
            if (string.Equals(sheetName, "DARIYER FILTER", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sheetName, "SWICH", StringComparison.OrdinalIgnoreCase))
            {
                return GetPartsForSheetVertical(ws, rows, cols);
            }

            // ---------------------------------------------------------------
            // Phase 1: detect whether this is a multi-box sheet.
            // Look through the first 6 rows for a cell containing "BOX" or "MASTER".
            // ---------------------------------------------------------------
            int headerRow = -1;
            for (int r = 1; r <= Math.Min(6, rows) && headerRow < 0; r++)
                for (int c = 1; c <= cols; c++)
                {
                    string t = ws.Cells[r, c].Text.Trim().ToUpper();
                    if (t.Contains("BOX") || t.Contains("MASTER")) { headerRow = r; break; }
                }

            if (headerRow < 0)
                return GetPartsForSheetVertical(ws, rows, cols);  // No box headers found

            // ---------------------------------------------------------------
            // Phase 2: collect box groups from the header row.
            //   "Data groups"  = first occurrence of each box label â†’ have QTY + SALE cols.
            //   "Price groups" = subsequent occurrences of the same label â†’ only PARTS NO + COST.
            // ---------------------------------------------------------------
            var seenLabels  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var dataGroups  = new List<(string Label, int StartCol)>();
            var priceGroups = new List<(string Label, int StartCol)>();

            for (int c = 1; c <= cols; c++)
            {
                string raw   = ws.Cells[headerRow, c].Text.Trim();
                string upper = raw.ToUpper();
                if (!upper.Contains("BOX") && !upper.Contains("MASTER")) continue;

                // Normalize "BOX -01" â†’ "BOX-01" for deduplication
                string norm = System.Text.RegularExpressions.Regex.Replace(raw, @"\s*-\s*", "-").Trim();

                if (!seenLabels.Contains(norm))
                {
                    seenLabels.Add(norm);
                    dataGroups.Add((raw, c));
                }
                else
                {
                    priceGroups.Add((raw, c));
                }
            }

            // Assign end columns (start of next group - 1)
            static List<(string L, int S, int E)> WithEnds(List<(string Label, int StartCol)> g, int totalCols)
            {
                var r = new List<(string, int, int)>(g.Count);
                for (int i = 0; i < g.Count; i++)
                    r.Add((g[i].Label, g[i].StartCol, i + 1 < g.Count ? g[i + 1].StartCol - 1 : totalCols));
                return r;
            }

            var dataGroupsFull  = WithEnds(dataGroups,  cols);
            var priceGroupsFull = WithEnds(priceGroups, cols);

            // ---------------------------------------------------------------
            // Phase 3: locate the sub-header row (row with PARTS NO / QTY / COST/PRICE).
            // ---------------------------------------------------------------
            int detailRow = headerRow + 1;
            for (int r = headerRow + 1; r <= Math.Min(headerRow + 4, rows); r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string t = ws.Cells[r, c].Text.Trim().ToUpper();
                    if (t.Contains("PART") || t.Contains("PATR") || t == "QTY" || t.Contains("PRICE"))
                    { detailRow = r; goto foundDetail; }
                }
            }
            foundDetail:
            int dataStartRow = detailRow + 1;

            // ---------------------------------------------------------------
            // Phase 4: build price map  (box-label, partNo) â†’ cost price
            //          from the right-hand "price groups".
            // ---------------------------------------------------------------
            // key = (normalized-box-label, normalized-partNo), value = cost price
            var priceMap = new Dictionary<(string, string), double>(
                comparer: new TupleIgnoreCaseComparer());

            foreach (var (label, startCol, endCol) in priceGroupsFull)
            {
                // In the price section the sub-header row has: PARTS NO | COST (QAR)
                int pPartCol = -1, pCostCol = -1;
                for (int c = startCol; c <= endCol; c++)
                {
                    string t = ws.Cells[detailRow, c].Text.Trim().ToUpper();
                    if (pPartCol < 0 && (t.Contains("PART") || t.Contains("PATR"))) pPartCol = c;
                    else if (pCostCol < 0 && (t.Contains("COST") || t.Contains("PRICE")))         pCostCol = c;
                }
                if (pPartCol < 0) pPartCol = startCol;
                if (pCostCol < 0 && startCol + 1 <= endCol) pCostCol = startCol + 1;

                // Normalize label (strip spaces around dash)
                string normLabel = System.Text.RegularExpressions.Regex.Replace(label, @"\s*-\s*", "-").Trim();

                for (int r = dataStartRow; r <= rows; r++)
                {
                    string pno = ws.Cells[r, pPartCol].Text.Trim();
                    if (string.IsNullOrEmpty(pno)) continue;
                    if (pCostCol > 0 && double.TryParse(ws.Cells[r, pCostCol].Text.Trim(), out double pc) && pc > 0)
                        priceMap[(normLabel, NormalizePart(pno))] = pc;
                }
            }

            // ---------------------------------------------------------------
            // Phase 5: iterate data groups and emit one entry per part per box.
            // ---------------------------------------------------------------
            foreach (var (label, startCol, endCol) in dataGroupsFull)
            {
                int partNoIdx = -1, qtyIdx = -1, salePriceIdx = -1, costPriceIdx = -1;

                for (int c = startCol; c <= endCol; c++)
                {
                    string t = ws.Cells[detailRow, c].Text.Trim().ToUpper();
                    if (string.IsNullOrEmpty(t)) continue;

                    if (partNoIdx < 0 && (t.Contains("PART") || t.Contains("PATR")))
                        partNoIdx = c;
                    else if (qtyIdx < 0 && (t == "QTY" || t == "QTV" || t.Contains("QTY")))
                        qtyIdx = c;
                    else if (salePriceIdx < 0 && (t == "SALE" || t == "SAL" || t.Contains("SALE PRICE") || t.Contains("SELL") || t.Contains("RETAIL")))
                        salePriceIdx = c;
                    else if (costPriceIdx < 0 && t.Contains("COST"))
                        costPriceIdx = c;
                }

                // Fallback column positions
                if (partNoIdx < 0) partNoIdx = startCol;
                if (qtyIdx < 0 && startCol + 1 <= endCol)   qtyIdx = startCol + 1;

                string normLabel = System.Text.RegularExpressions.Regex.Replace(label, @"\s*-\s*", "-").Trim();

                for (int r = dataStartRow; r <= rows; r++)
                {
                    string partNo = ws.Cells[r, partNoIdx].Text.Trim();
                    if (string.IsNullOrEmpty(partNo)) continue;
                    string upperPN = partNo.ToUpper();
                    if (IsStockHeaderOrCategory(upperPN)) continue;

                    double qty = 0, sale = 0, cost = 0;
                    if (qtyIdx > 0)       double.TryParse(CleanQtyText(ws.Cells[r, qtyIdx].Text), out qty);
                    if (salePriceIdx > 0) double.TryParse(ws.Cells[r, salePriceIdx].Text.Trim(), out sale);
                    if (costPriceIdx > 0) double.TryParse(ws.Cells[r, costPriceIdx].Text.Trim(), out cost);

                    // Try price map (from right-hand price section)
                    if (cost <= 0 && priceMap.TryGetValue((normLabel, NormalizePart(partNo)), out double mapped))
                        cost = mapped;
                    if (sale <= 0 && cost > 0) sale = cost;

                    results.Add(new StockPartEntry
                    {
                        PartNo    = partNo,
                        BoxLabel  = label,
                        Qty       = qty,
                        SalePrice = sale,
                        CostPrice = cost
                    });
                }
            }

            return results;
        }

        // ---------------------------------------------------------------
        // Helper: normalise a part-number string for dictionary lookup.
        // Strips internal spaces and uppercases.
        // ---------------------------------------------------------------
        private static string NormalizePart(string p) =>
            System.Text.RegularExpressions.Regex.Replace(p.Trim().ToUpper(), @"\s+", " ");

        // Comparer for (string, string) tuple ignoring case
        private class TupleIgnoreCaseComparer : IEqualityComparer<(string, string)>
        {
            public bool Equals((string, string) x, (string, string) y) =>
                string.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(x.Item2, y.Item2, StringComparison.OrdinalIgnoreCase);

            public int GetHashCode((string, string) obj) =>
                HashCode.Combine(
                    obj.Item1?.ToUpperInvariant()?.GetHashCode() ?? 0,
                    obj.Item2?.ToUpperInvariant()?.GetHashCode() ?? 0);
        }

        /// <summary>
        /// Handles all vertical-layout sheets (one part per row).
        /// Correctly maps SALE/SAL as sale-price (not seal/adjustment),
        /// and recognises every sheet-specific part-column heading:
        /// EVAPORATOR, HEATER, AC COMPRESSUER, CONDANSOER, SWICH, FILTER, FAN,
        /// PARTS, PART, PATR.
        /// </summary>
        private static List<StockPartEntry> GetPartsForSheetVertical(ExcelWorksheet ws, int rows, int cols)
        {
            var results = new List<StockPartEntry>();

            // Indices (1-based columns)
            int startRow     = 2;   // default: row 2 is data start
            int partNoIdx    = 1;   // default: column 1
            int qtyIdx       = 2;   // default: column 2
            int salePriceIdx = -1;
            int costPriceIdx = -1;
            int sealAdjIdx   = -1;  // adjustment/seal column (negative deltas like -1, -2â€¦)
            int modelIdx     = -1;

            // Scan up to 3 header rows (handles 2-row headers like AC COMORESSURE)
            bool foundHeader = false;
            for (int r = 1; r <= Math.Min(3, rows) && !foundHeader; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string raw = ws.Cells[r, c].Text.Trim();
                    string t   = raw.ToUpper();
                    if (string.IsNullOrEmpty(t)) continue;

                    // â”€â”€ Part-number column â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    // Matches any known sheet-name used as a header, plus generic PART/PATR
                    bool isPartHeader =
                        t == "PARTS NO" || t == "PARTS" || t == "PART NO" ||
                        t.Contains("PATR") ||
                        t == "EVAPORATOR" || t.StartsWith("EVAPORATOR") ||
                        t == "HEATER"     || t.StartsWith("HEATER") ||
                        t == "FAN"        ||
                        t == "CONDANSOER" || t.StartsWith("CONDANSOER") ||
                        t == "SWICH"      ||
                        t == "FILTER"     ||
                        t.Contains("AC COMPRESSUER") || t.Contains("AC COMP");

                    if (partNoIdx == 1 /* still default */ && isPartHeader)
                    {
                        partNoIdx = c;
                        startRow  = r + 1;
                        foundHeader = true;
                    }

                    // â”€â”€ Qty column â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if (qtyIdx == 2 /* still default */ &&
                        (t == "QTY" || t == "QTV" || t.Contains("QTY")))
                        qtyIdx = c;

                    // â”€â”€ Sale price â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    // Check SALE PRICE or standalone SALE / SAL *before* SEAL
                    if (salePriceIdx < 0 &&
                        (t == "SALE" || t == "SAL" ||
                         t.Contains("SALE PRICE") || t.Contains("SELL") || t.Contains("RETAIL")))
                        salePriceIdx = c;

                    // â”€â”€ Seal / adjustment column â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    // Only "SEAL" (not standalone "SALE") qualifies
                    if (sealAdjIdx < 0 && t.Contains("SEAL") && !t.Contains("SALE"))
                        sealAdjIdx = c;

                    // â”€â”€ Cost price â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if (costPriceIdx < 0 && t.Contains("COST"))
                        costPriceIdx = c;

                    // â”€â”€ Model / brand â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
                    if (modelIdx < 0 && (t.Contains("MODEL") || t.Contains("BRAND")))
                        modelIdx = c;
                }
            }

            // â”€â”€ Data rows â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
            for (int r = startRow; r <= rows; r++)
            {
                string partNo = ws.Cells[r, partNoIdx].Text.Trim();
                if (string.IsNullOrEmpty(partNo)) continue;

                string upperPN = partNo.ToUpper();
                // Skip header-repeat rows or summary rows
                if (IsStockHeaderOrCategory(upperPN))
                    continue;

                double qty = 0, sale = 0, cost = 0;
                if (double.TryParse(CleanQtyText(ws.Cells[r, qtyIdx].Text), out double rawQty))
                    qty = rawQty;

                // SALE column: may contain negative values (adjustments like -2)
                // Only use positive values as invoice price
                if (salePriceIdx > 0)
                {
                    double.TryParse(ws.Cells[r, salePriceIdx].Text.Trim(), out sale);
                    if (sale < 0) sale = 0; // negative = adjustment indicator, not a price
                }

                if (costPriceIdx > 0)
                    double.TryParse(ws.Cells[r, costPriceIdx].Text.Trim(), out cost);

                if (sale <= 0 && cost > 0) sale = cost;

                string model = modelIdx > 0 ? ws.Cells[r, modelIdx].Text.Trim() : string.Empty;

                results.Add(new StockPartEntry
                {
                    PartNo    = partNo,
                    BoxLabel  = string.Empty,
                    Model     = model,
                    Qty       = qty,
                    SalePrice = sale,
                    CostPrice = cost
                });
            }

            return results;
        }

        // ----------------------------------------------------------------
        // 3. Full parse for Inventory Sync Window
        // ----------------------------------------------------------------
        public static (List<ParsedPartInfo> ExistingParts, List<ParsedPartInfo> NewParts) ParseExcelFile(string filePath)
        {
            var existingParts = new List<ParsedPartInfo>();
            var newParts = new List<ParsedPartInfo>();

            var existingPartNos = DatabaseHelper
                .ExecuteQuery("SELECT PartNo FROM Parts;", reader => reader["PartNo"].ToString()!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var parsedPartsDict = new Dictionary<string, ParsedPartInfo>(StringComparer.OrdinalIgnoreCase);

            using var pkg = OpenProtected(filePath);

            foreach (var ws in pkg.Workbook.Worksheets)
            {
                string sheetName = ws.Name.Trim();
                
                string category;
                string prefix;
                if (WorksheetSchema.TryGetValue(sheetName, out var mapping))
                {
                    category = mapping.Category;
                    prefix = mapping.Prefix;
                }
                else
                {
                    category = sheetName;
                    prefix = sheetName.Length >= 2 ? sheetName.Substring(0, 2).ToUpper() : sheetName.ToUpper();
                }

                int rows = ws.Dimension?.Rows ?? 0;
                int cols = ws.Dimension?.Columns ?? 0;
                if (rows < 1) continue;

                int startRow = 1;
                int partNoIdx = -1;
                int qtyIdx = -1;
                int sealIdx = -1;
                int locationIdx = -1;
                int modelIdx = -1;
                int costPriceIdx = -1;
                int salePriceIdx = -1;

                // Ø§Ù„ÙØ­Øµ Ø§Ù„Ø°ÙƒÙŠ ÙÙŠ Ø£ÙˆÙ„ 5 ØµÙÙˆÙ
                for (int r = 1; r <= Math.Min(5, rows); r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        string txt = ws.Cells[r, c].Text.Trim().ToUpper();
                        if (string.IsNullOrEmpty(txt)) continue;

                        if (partNoIdx == -1 && (txt.Contains("PART") || txt.Contains("PATR") || txt.Contains("EVAPORATOR") || txt.Contains("FAN") || txt.Contains("AC COMPRESSUER")))
                        {
                            partNoIdx = c;
                            startRow = r + 1;
                        }
                        if (qtyIdx == -1 && (txt == "QTY" || txt == "QTV" || txt.Contains("QTY")))
                        {
                            qtyIdx = c;
                        }
                        if (sealIdx == -1 && (txt.Contains("SEAL") || txt.Contains("SALE")))
                        {
                            sealIdx = c;
                        }
                        if (locationIdx == -1 && (txt.Contains("LOCATION") || txt.Contains("PLAS")))
                        {
                            locationIdx = c;
                        }
                        if (modelIdx == -1 && (txt.Contains("MODEL") || txt.Contains("BRAND")))
                        {
                            modelIdx = c;
                        }
                        if (costPriceIdx == -1 && (txt.Contains("COST PRICE") || txt.Contains("COST")))
                        {
                            costPriceIdx = c;
                        }
                        if (salePriceIdx == -1 && (txt.Contains("SALE PRICE") || txt.Contains("SELL")))
                        {
                            salePriceIdx = c;
                        }
                    }
                    if (partNoIdx != -1 && qtyIdx != -1) break;
                }

                if (partNoIdx <= 0) partNoIdx = 1;
                if (qtyIdx <= 0) qtyIdx = 2;

                for (int r = startRow; r <= rows; r++)
                {
                    string partNo = ws.Cells[r, partNoIdx].Text.Trim();
                    if (string.IsNullOrEmpty(partNo) || partNo.ToUpper().Contains("PART") || partNo.ToUpper().Contains("QTY"))
                        continue;

                    double qty = 0, seal = 0, cost = 0, sale = 0;
                    double.TryParse(CleanQtyText(ws.Cells[r, qtyIdx].Text), out qty);
                    if (sealIdx > 0) double.TryParse(CleanQtyText(ws.Cells[r, sealIdx].Text), out seal);
                    if (costPriceIdx > 0) double.TryParse(ws.Cells[r, costPriceIdx].Text.Trim(), out cost);
                    if (salePriceIdx > 0) double.TryParse(ws.Cells[r, salePriceIdx].Text.Trim(), out sale);

                    // âœ¨ ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ø­Ù…Ø§ÙŠØ© Ù‡Ù†Ø§ Ø£ÙŠØ¶Ø§Ù‹ Ù„Ø¹Ù…Ù„ÙŠØ© Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© ÙˆØ§Ù„Ù…Ø·Ø§Ø¨Ù‚Ø© Ø§Ù„ÙƒØ§Ù…Ù„Ø© Ù…Ø¹ Ù‚Ø§Ø¹Ø¯Ø© Ø§Ù„Ø¨ÙŠØ§Ù†Ø§Øª
                    if (sale == 0 && cost > 0)
                    {
                        sale = cost;
                    }

                    string location = locationIdx > 0 ? ws.Cells[r, locationIdx].Text.Trim() : string.Empty;
                    string model = modelIdx > 0 ? ws.Cells[r, modelIdx].Text.Trim() : string.Empty;

                    if (parsedPartsDict.TryGetValue(partNo, out var existingParsed))
                    {
                        existingParsed.QtyInExcel += qty;
                        existingParsed.SealValue += seal;
                        if (!string.IsNullOrEmpty(location)) existingParsed.Location = location;
                        if (!string.IsNullOrEmpty(model)) existingParsed.CarModel = model;
                        if (cost > 0) existingParsed.CostPrice = cost;
                        if (sale > 0) existingParsed.SalePrice = sale;
                    }
                    else
                    {
                        var info = new ParsedPartInfo
                        {
                            PartNo = partNo,
                            Description = $"{category} â€” {partNo}",
                            Category = category,
                            Prefix = prefix,
                            QtyInExcel = qty,
                            SealValue = seal,
                            Location = location,
                            CarModel = model,
                            CostPrice = cost,
                            SalePrice = sale,
                            IsNew = !existingPartNos.Contains(partNo)
                        };
                        parsedPartsDict[partNo] = info;
                    }
                }
            }

            foreach (var part in parsedPartsDict.Values)
            {
                if (part.IsNew)
                    newParts.Add(part);
                else
                    existingParts.Add(part);
            }

            return (existingParts, newParts);
        }

        // ----------------------------------------------------------------
        // 4. Write-back stock deductions to Excel after invoice is locked
        // ----------------------------------------------------------------
        /// <summary>
        /// Deducts sold quantities from the QTY column in the original Excel file on disk.
        /// Reads via FileShare.ReadWrite (works even when Excel has the file open for viewing),
        /// modifies in-memory, then writes the result back to the original file path.
        /// Supports both vertical sheets (one row per part) and multi-box sheets (HOUSE).
        /// </summary>
        public static void WriteBackStockDeductions(string filePath, Dictionary<string, double> partDeductions)
        {
            if (partDeductions == null || partDeductions.Count == 0) return;

            // Step 1: Read the current file into a MemoryStream.
            // FileShare.ReadWrite allows reading even when Excel has it open.
            MemoryStream inputStream;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                inputStream = new MemoryStream((int)fs.Length);
                fs.CopyTo(inputStream);
                inputStream.Position = 0;
            }

            var remainingDeductions = partDeductions
                .Where(d => !string.IsNullOrWhiteSpace(d.Key) && d.Value > 0)
                .GroupBy(d => d.Key.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Sum(d => d.Value), StringComparer.OrdinalIgnoreCase);

            // Step 2: Open the workbook from memory.
            ExcelPackage? pkg = null;
            bool isEncrypted = false;
            try
            {
                pkg = new ExcelPackage(inputStream);
                _ = pkg.Workbook.Worksheets.Count;
            }
            catch
            {
                pkg?.Dispose();
                inputStream.Position = 0;
                pkg = new ExcelPackage(inputStream, ExcelPassword);
                isEncrypted = true;
            }

            try
            {
                // Step 3: Walk every sheet and deduct from matching QTY cells.
                foreach (var ws in pkg.Workbook.Worksheets)
                {
                    int rows = ws.Dimension?.Rows ?? 0;
                    int cols = ws.Dimension?.Columns ?? 0;
                    if (rows < 1) continue;

                    int startRow  = 1;
                    int partNoCol = -1;
                    var qtyCols   = new List<int>();

                    for (int r = 1; r <= Math.Min(6, rows); r++)
                    {
                        bool rowHasPartHeader = false;
                        for (int c = 1; c <= cols; c++)
                        {
                            string txt = ws.Cells[r, c].Text.Trim().ToUpper();
                            if (string.IsNullOrEmpty(txt)) continue;

                            if (partNoCol < 0 &&
                                (txt.Contains("PART") || txt.Contains("PATR") ||
                                 txt == "EVAPORATOR"  || txt.StartsWith("EVAPORATOR") ||
                                 txt == "HEATER"      || txt.StartsWith("HEATER") ||
                                 txt == "FAN"         || txt == "SWICH" || txt == "FILTER" ||
                                 txt.Contains("CONDANSOER") || txt.Contains("AC COMP")))
                            {
                                partNoCol        = c;
                                startRow         = r + 1;
                                rowHasPartHeader = true;
                            }

                            if (txt == "QTY" || txt == "QTV" || txt.Contains("QTY"))
                                if (!qtyCols.Contains(c)) qtyCols.Add(c);
                        }

                        if (rowHasPartHeader) break;
                    }

                    if (partNoCol <= 0) partNoCol = 1;
                    if (qtyCols.Count == 0)
                    {
                        if (cols >= 2) qtyCols.Add(2);
                        else continue;
                    }

                    for (int r = startRow; r <= rows; r++)
                    {
                        string partNo = ws.Cells[r, partNoCol].Text.Trim();
                        if (string.IsNullOrEmpty(partNo)) continue;

                        string upperPN = partNo.ToUpper();
                        if (IsStockHeaderOrCategory(upperPN))
                            continue;

                        if (!remainingDeductions.TryGetValue(partNo, out double remainingQty) || remainingQty <= 0)
                            continue;

                        // Deduct from the first QTY column with a positive value.
                        foreach (int qtyCol in qtyCols)
                        {
                            string rawCell = ws.Cells[r, qtyCol].Text.Trim();
                            if (string.IsNullOrEmpty(rawCell)) continue;

                            if (double.TryParse(CleanQtyText(rawCell), out double currentQty) && currentQty > 0)
                            {
                                double deductedQty = Math.Min(currentQty, remainingQty);
                                double newQty = currentQty - deductedQty;
                                ws.Cells[r, qtyCol].Value = newQty;
                                remainingDeductions[partNo] -= deductedQty;
                                break;
                            }
                        }
                    }
                }

                // Step 4: Serialize modified workbook then overwrite the file on disk.
                using var outputStream = new MemoryStream();
                if (isEncrypted)
                    pkg.SaveAs(outputStream, ExcelPassword);
                else
                    pkg.SaveAs(outputStream);

                outputStream.Position = 0;

                WriteWorkbookBytesWithRetry(filePath, outputStream.ToArray());
            }
            finally
            {
                pkg?.Dispose();
                inputStream.Dispose();
            }
        }

        // ----------------------------------------------------------------
        // 4b. Write-back stock returns to Excel after part return
        // ----------------------------------------------------------------
        /// <summary>
        /// Restores returned quantities back to the QTY column in the original Excel file on disk.
        /// Reads via FileShare.ReadWrite, modifies in-memory, then writes back.
        /// Supports both vertical and multi-box layouts.
        /// </summary>
        public static void WriteBackStockReturns(string filePath, Dictionary<string, double> partReturns)
        {
            if (partReturns == null || partReturns.Count == 0) return;

            MemoryStream inputStream;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                inputStream = new MemoryStream((int)fs.Length);
                fs.CopyTo(inputStream);
                inputStream.Position = 0;
            }

            var returnMap = new Dictionary<string, double>(partReturns, StringComparer.OrdinalIgnoreCase);

            ExcelPackage? pkg = null;
            bool isEncrypted = false;
            try
            {
                pkg = new ExcelPackage(inputStream);
                _ = pkg.Workbook.Worksheets.Count;
            }
            catch
            {
                pkg?.Dispose();
                inputStream.Position = 0;
                pkg = new ExcelPackage(inputStream, ExcelPassword);
                isEncrypted = true;
            }

            try
            {
                foreach (var ws in pkg.Workbook.Worksheets)
                {
                    int rows = ws.Dimension?.Rows ?? 0;
                    int cols = ws.Dimension?.Columns ?? 0;
                    if (rows < 1) continue;

                    int startRow  = 1;
                    int partNoCol = -1;
                    var qtyCols   = new List<int>();

                    for (int r = 1; r <= Math.Min(6, rows); r++)
                    {
                        bool rowHasPartHeader = false;
                        for (int c = 1; c <= cols; c++)
                        {
                            string txt = ws.Cells[r, c].Text.Trim().ToUpper();
                            if (string.IsNullOrEmpty(txt)) continue;

                            if (partNoCol < 0 &&
                                (txt.Contains("PART") || txt.Contains("PATR") ||
                                 txt == "EVAPORATOR"  || txt.StartsWith("EVAPORATOR") ||
                                 txt == "HEATER"      || txt.StartsWith("HEATER") ||
                                 txt == "FAN"         || txt == "SWICH" || txt == "FILTER" ||
                                 txt.Contains("CONDANSOER") || txt.Contains("AC COMP")))
                            {
                                partNoCol        = c;
                                startRow         = r + 1;
                                rowHasPartHeader = true;
                            }

                            if (txt == "QTY" || txt == "QTV" || txt.Contains("QTY"))
                                if (!qtyCols.Contains(c)) qtyCols.Add(c);
                        }

                        if (rowHasPartHeader) break;
                    }

                    if (partNoCol <= 0) partNoCol = 1;
                    if (qtyCols.Count == 0)
                    {
                        if (cols >= 2) qtyCols.Add(2);
                        else continue;
                    }

                    for (int r = startRow; r <= rows; r++)
                    {
                        string partNo = ws.Cells[r, partNoCol].Text.Trim();
                        if (string.IsNullOrEmpty(partNo)) continue;

                        string upperPN = partNo.ToUpper();
                        if (IsStockHeaderOrCategory(upperPN))
                            continue;

                        if (!returnMap.TryGetValue(partNo, out double returnedQty)) continue;

                        foreach (int qtyCol in qtyCols)
                        {
                            string rawCell = ws.Cells[r, qtyCol].Text.Trim();
                            double currentQty = 0;
                            if (!string.IsNullOrEmpty(rawCell))
                            {
                                double.TryParse(CleanQtyText(rawCell), out currentQty);
                            }
                            double newQty = currentQty + returnedQty;
                            ws.Cells[r, qtyCol].Value = newQty;
                            break;
                        }
                    }
                }

                using var outputStream = new MemoryStream();
                if (isEncrypted)
                    pkg.SaveAs(outputStream, ExcelPassword);
                else
                    pkg.SaveAs(outputStream);

                outputStream.Position = 0;

                using var outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                outputStream.CopyTo(outFile);
            }
            finally
            {
                pkg?.Dispose();
                inputStream.Dispose();
            }
        }

        // ----------------------------------------------------------------
        // 5. Commit parsed parts to the SQLite database (bulk upsert)
        // ----------------------------------------------------------------
        public static void CommitSyncToDatabase(
            List<ParsedPartInfo> existingParts,
            List<ParsedPartInfo> newParts)
        {
            DatabaseHelper.ExecuteTransaction((conn, trans) =>
            {
                foreach (var part in existingParts)
                {
                    string update = @"
                        UPDATE Parts
                        SET QtyInStock   = @qty,
                            CostPrice    = @cost,
                            SalePrice    = @sale,
                            Location     = @loc,
                            CarModel     = @model,
                            LastSyncedAt = @synced
                        WHERE PartNo = @partNo;";

                    using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(update, conn, trans);
                    cmd.Parameters.AddWithValue("@qty", part.CalculatedQty);
                    cmd.Parameters.AddWithValue("@cost", part.CostPrice);
                    cmd.Parameters.AddWithValue("@sale", part.SalePrice);
                    cmd.Parameters.AddWithValue("@loc", part.Location);
                    cmd.Parameters.AddWithValue("@model", part.CarModel);
                    cmd.Parameters.AddWithValue("@synced", DateTime.UtcNow.ToString("o"));
                    cmd.Parameters.AddWithValue("@partNo", part.PartNo);
                    cmd.ExecuteNonQuery();
                }

                foreach (var part in newParts)
                {
                    string insert = @"
                        INSERT INTO Parts (PartNo, Description, Category, Prefix, QtyInStock, CostPrice, SalePrice, Location, CarModel, LastSyncedAt, LastUpdatedAt)
                        VALUES (@partNo, @desc, @cat, @prefix, @qty, @cost, @sale, @loc, @model, @synced, @synced);";

                    using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(insert, conn, trans);
                    cmd.Parameters.AddWithValue("@partNo", part.PartNo);
                    cmd.Parameters.AddWithValue("@desc", part.Description);
                    cmd.Parameters.AddWithValue("@cat", part.Category);
                    cmd.Parameters.AddWithValue("@prefix", part.Prefix);
                    cmd.Parameters.AddWithValue("@qty", part.CalculatedQty);
                    cmd.Parameters.AddWithValue("@cost", part.CostPrice);
                    cmd.Parameters.AddWithValue("@sale", part.SalePrice);
                    cmd.Parameters.AddWithValue("@loc", part.Location);
                    cmd.Parameters.AddWithValue("@model", part.CarModel);
                    cmd.Parameters.AddWithValue("@synced", DateTime.UtcNow.ToString("o"));
                    cmd.ExecuteNonQuery();
                }
            });
        }

        /// <summary>
        /// Replaces the workbook content while allowing another process (such as Excel) to
        /// keep a shared handle open. Excel can briefly hold the file while autosaving, so
        /// retry a few times before reporting a real synchronization failure.
        /// </summary>
        private static void WriteWorkbookBytesWithRetry(string filePath, byte[] workbookBytes)
        {
            const int maxAttempts = 4;
            Exception? lastError = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    using var outFile = new FileStream(filePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
                    outFile.SetLength(0);
                    outFile.Write(workbookBytes, 0, workbookBytes.Length);
                    outFile.Flush(flushToDisk: true);
                    return;
                }
                catch (IOException ex)
                {
                    lastError = ex;
                }
                catch (UnauthorizedAccessException ex)
                {
                    lastError = ex;
                }

                if (attempt < maxAttempts)
                    System.Threading.Thread.Sleep(attempt * 300);
            }

            throw new IOException(
                "The stock workbook is still locked after several write attempts. " +
                "Please save the workbook in Excel and try again.", lastError);
        }

        private static ExcelPackage OpenProtected(string filePath)
        {
            // Copy entire file into a MemoryStream first so EPPlus never tries to lock
            // the file on disk â€” this allows the file to be open in Excel simultaneously.
            // We use a FileStream with FileShare.ReadWrite so Windows does not block us
            // even if another process (e.g. Excel) has the file open.
            MemoryStream fileStream;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                fileStream = new MemoryStream();
                fs.CopyTo(fileStream);
                fileStream.Position = 0;
            }

            ExcelPackage? pkg = null;
            try
            {
                pkg = new ExcelPackage(fileStream);
                // Force EPPlus to parse the workbook structure (catches encryption errors early)
                _ = pkg.Workbook.Worksheets.Count;
                return pkg;
            }
            catch
            {
                if (pkg != null)
                {
                    try { pkg.Dispose(); } catch { }
                }
                // Fallback: re-open from the same bytes with the default password ("0000")
                fileStream.Position = 0;
                return new ExcelPackage(fileStream, ExcelPassword);
            }
        }

        public static async Task StreamImportPartsAsync(IEnumerable<ParsedPartInfo> parts, IProgress<int>? progress = null)
        {
            var channel = Channel.CreateUnbounded<ParsedPartInfo>(new UnboundedChannelOptions { SingleReader = true });

            var producer = Task.Run(async () =>
            {
                foreach (var part in parts)
                {
                    if (part.IsSelectedToImport)
                    {
                        await channel.Writer.WriteAsync(part);
                    }
                }
                channel.Writer.Complete();
            });

            var consumer = Task.Run(async () =>
            {
                int count = 0;
                while (await channel.Reader.WaitToReadAsync())
                {
                    while (channel.Reader.TryRead(out var part))
                    {
                        DatabaseHelper.ExecuteNonQuery(@"
                            INSERT INTO Parts (PartNo, Description, Category, Prefix, QtyInStock, CostPrice, SalePrice, Location, CarModel, LastSyncedAt, LastUpdatedAt)
                            VALUES (@pno, @desc, @cat, @pref, @qty, @cost, @sale, @loc, @model, @sync, @upd)
                            ON CONFLICT(PartNo) DO UPDATE SET
                                QtyInStock = excluded.QtyInStock,
                                CostPrice = excluded.CostPrice,
                                SalePrice = excluded.SalePrice,
                                LastSyncedAt = excluded.LastSyncedAt;",
                            new Microsoft.Data.Sqlite.SqliteParameter("@pno", part.PartNo),
                            new Microsoft.Data.Sqlite.SqliteParameter("@desc", part.Description),
                            new Microsoft.Data.Sqlite.SqliteParameter("@cat", part.Category),
                            new Microsoft.Data.Sqlite.SqliteParameter("@pref", part.Prefix),
                            new Microsoft.Data.Sqlite.SqliteParameter("@qty", part.CalculatedQty),
                            new Microsoft.Data.Sqlite.SqliteParameter("@cost", part.CostPrice),
                            new Microsoft.Data.Sqlite.SqliteParameter("@sale", part.SalePrice),
                            new Microsoft.Data.Sqlite.SqliteParameter("@loc", part.Location),
                            new Microsoft.Data.Sqlite.SqliteParameter("@model", part.CarModel),
                            new Microsoft.Data.Sqlite.SqliteParameter("@sync", DateTime.UtcNow.ToString("o")),
                            new Microsoft.Data.Sqlite.SqliteParameter("@upd", DateTime.UtcNow.ToString("o"))
                        );
                        count++;
                        if (count % 100 == 0) progress?.Report(count);
                    }
                }
                progress?.Report(count);
            });

            await Task.WhenAll(producer, consumer);
        }
    }
}
