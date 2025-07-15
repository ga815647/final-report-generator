using System;
using System.Windows.Forms;

namespace MyApp.Utils
{
    public static class FileSelector
    {
        /// <summary>
        /// 顯示單一對話框，同時輸入民國年、月並選擇檔案。
        /// 如取消或關閉，顯示訊息後結束程式。
        /// </summary>
        public static (string FilePath, int Year, int Month) SelectExcelFileWithDate()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using var form = new DateFileSelectorForm();
            if (form.ShowDialog() != DialogResult.OK)
            {
                Console.WriteLine("操作已取消，程式結束。");
                Environment.Exit(0);
            }

            return (form.FilePath, form.Year, form.Month);
        }
    }
}
