using System;
using System.Collections.Generic;
using System.Text;
using MyApp.Utils;

namespace MyApp
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ConfigureConsoleEncoding();

            // 1. 取得使用者選擇的檔案路徑 + 年 + 月
            var (filePath, year, month) = FileSelector.SelectExcelFileWithDate();
            Console.WriteLine($"選擇檔案：{filePath}，年度：{year}；月份：{month}");

            // 2. 讀取 Excel、捕捉可能拋出的錯誤
            List<string> values;
            try
            {
                values = ExcelReader.ReadMainAccountingOffice(filePath, year, month);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"處理 Excel 時發生錯誤：{ex.Message}");
                return;    // 讀不到資料就結束
            }

            // 3. 產生 Word 報表、捕捉可能拋出的錯誤
            try
            {
                ReportWriter.WriteReport(year, month, values);
                Console.WriteLine("Word 報表已輸出到桌面。");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"匯出 Word 時發生錯誤：{ex.Message}");
            }
        }

        /// <summary>
        /// 封裝一次性設定：避免亂碼
        /// </summary>
        private static void ConfigureConsoleEncoding()
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding  = Encoding.UTF8;
        }
    }
}
