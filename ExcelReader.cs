using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace MyApp.Utils
{
    public static class ExcelReader
    {
        private const string LicenseOwner = "My Name";

        // 全域儲存主計室 + 附件 會計的原始儲存格字串
        private static readonly List<string> rawCellTexts = new List<string>();

        public static List<string> ReadMainAccountingOffice(string filePath, int year, int month)
        {
            ValidateFilePath(filePath);
            using var package = CreatePackage(filePath);

            rawCellTexts.Clear();

            // 1. 讀「主計室」固定儲存格
            ReadFixedCells(package,
                        "主計室",
                        new[] { "B7","C7","D7","F7","H7","I7","J7","K7","L7","N7" });

            // 2. 讀「附件 會計」與「藥衛材(不含花榮)」動態列
            MergeMainAndAnnex(package,
                            "附件 會計",
                            year,
                            month,
                            new[] { "I","B","L","E","N","J","C","M","F","O","H","K","P" });
            MergeMainAndAnnex(package,
                            "藥衛材(不含花榮)",
                            year,
                            month,
                            new[] { "D","K","N","G","Q","H","R" });

            // 3. 格式化主／附件資料
            var formattedCellTexts = FormatRawValues(rawCellTexts);

            // 4. 讀三個工作表並附加值
            AppendSheetValue(package, formattedCellTexts, "藥品", month, 7);
            AppendSheetValue(package, formattedCellTexts, "衛材", month, 7);
            AppendSheetValue(package, formattedCellTexts, "藥衛", month, 7);
            AppendSheetValue(package, formattedCellTexts, "藥衛", month, 9);

            return formattedCellTexts;
        }


        /// <summary>
        /// 在指定工作表第2列尋找 "{month}月" 標題，取第7列數值並格式化兩位小數加入 list。
        /// </summary>
        private static void AppendSheetValue(ExcelPackage pkg,
                                            List<string> list,
                                            string sheetName,
                                            int month,
                                            int valueRow)
        {
            var ws = GetWorksheet(pkg, sheetName);
            var headerText = $"{month}月";

            int targetCol = -1;
            for (int c = ws.Dimension.Start.Column; c <= ws.Dimension.End.Column; c++)
            {
                if (ws.Cells[2, c].Text == headerText)
                {
                    targetCol = c;
                    break;
                }
            }

            if (targetCol < 0)
                throw new InvalidOperationException(
                    $"在工作表「{sheetName}」第2列找不到「{headerText}」。");

            var raw = ws.Cells[valueRow, targetCol].Text;
            if (!double.TryParse(raw, out double val))
                val = 0;

            // 固定格式：小數點兩位
            list.Add(val.ToString("F2"));
        }

        private static void ValidateFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new ArgumentException("無效的檔案路徑。", nameof(path));
        }

        private static ExcelPackage CreatePackage(string filePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal(LicenseOwner);
            return new ExcelPackage(new FileInfo(filePath));
        }

        private static ExcelWorksheet GetWorksheet(ExcelPackage pkg, string sheetName)
        {
            var ws = pkg.Workbook.Worksheets[sheetName];
            if (ws == null)
                throw new InvalidOperationException($"找不到工作表「{sheetName}」。");
            return ws;
        }

        private static void ReadFixedCells(ExcelPackage pkg,
                                           string sheetName,
                                           string[] addresses)
        {
            var ws = GetWorksheet(pkg, sheetName);
            foreach (var addr in addresses)
            {
                rawCellTexts.Add(ws.Cells[addr].Text);
            }
        }

        private static void MergeMainAndAnnex(ExcelPackage pkg,
                                              string annexSheet,
                                              int year,
                                              int month,
                                              string[] annexCols)
        {
            var wsAnnex = GetWorksheet(pkg, annexSheet);
            int row = FindRowByYearMonth(wsAnnex, year, month);

            foreach (var col in annexCols)
            {
                rawCellTexts.Add(wsAnnex.Cells[$"{col}{row}"].Text);
            }
        }

        private static int FindRowByYearMonth(ExcelWorksheet ws, int year, int month)
        {
            var target = $"{year:D3}/{month:D2}";
            for (int r = ws.Dimension.Start.Row; r <= ws.Dimension.End.Row; r++)
            {
                if (ws.Cells[r, 1].Text == target)
                    return r;
            }
            throw new InvalidOperationException(
                $"在工作表「{ws.Name}」的 A 欄找不到「{target}」。");
        }

        private static List<string> FormatRawValues(IEnumerable<string> rawList)
        {
            var result = new List<string>();
            foreach (var raw in rawList)
            {
                if (string.IsNullOrEmpty(raw))
                {
                    result.Add("0");
                }
                else if (double.TryParse(raw, out double val))
                {
                    // 判斷整數 vs 小數
                    if (Math.Abs(val - Math.Truncate(val)) < double.Epsilon)
                    {
                        // 整數：每三位加逗號
                        result.Add(((int)val).ToString("#,##0"));
                    }
                    else
                    {
                        // 小數：百分比 (兩位小數)
                        result.Add(val.ToString("P2"));
                    }
                }
                else
                {
                    result.Add(raw);
                }
            }
            return result;
        }

    }
}
