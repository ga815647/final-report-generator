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

            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // 捕捉任何從 Run() 垂直傳上來的未處理例外
                Console.Error.WriteLine($"發生未預期錯誤：{ex.Message}");
            }
            finally
            {
                // 無論正常結束、catch 裡 return，或在 Run() 內拋出例外，
                // finally 區塊裡的這兩行一定會被執行
                Console.WriteLine("按任意鍵結束…");
                Console.ReadKey(true);
            }
        }

        /// <summary>
        /// 把原本 Main 的邏輯都搬到這裡，方便統一在 Main 裡做 try/finally
        /// </summary>
        private static void Run()
        {
            string filePath;
            int year, month;

            // 1. 取得使用者選擇的檔案路徑 + 年 + 月
            try
            {
                (filePath, year, month) = FileSelector.SelectExcelFileWithDate();
                Console.WriteLine($"選擇檔案：{filePath}，年度：{year}；月份：{month}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"輸入資料錯誤：{ex.Message}");
                return;
            }

            // 2. 讀取 Excel
            List<string> values;
            try
            {
                values = ExcelReader.ReadMainAccountingOffice(filePath, year, month);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"處理 Excel 時發生錯誤：{ex.Message}");
                return;
            }

            // 3. 產生 Word 報表
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

        private static void ConfigureConsoleEncoding()
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding  = Encoding.UTF8;
        }
    }
}
