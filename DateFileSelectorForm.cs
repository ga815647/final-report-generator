using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MyApp.Utils
{
    public class DateFileSelectorForm : Form
    {
        private NumericUpDown nudYear;
        private NumericUpDown nudMonth;
        private TextBox       txtPath;
        private Button        btnBrowse;
        private Button        btnOK;
        private Button        btnCancel;

        public int Year     { get; private set; }
        public int Month    { get; private set; }
        public string FilePath { get; private set; } = string.Empty;

        public DateFileSelectorForm()
        {
            // 一律 new 控制項，避免 CS8618
            nudYear   = new NumericUpDown();
            nudMonth  = new NumericUpDown();
            txtPath   = new TextBox();
            btnBrowse = new Button();
            btnOK     = new Button();
            btnCancel = new Button();

            // Form 基本設定：改寬到 700px
            Text               = "請輸入年度、月份並選擇檔案";
            FormBorderStyle    = FormBorderStyle.FixedDialog;
            MaximizeBox        = false;
            MinimizeBox        = false;
            StartPosition      = FormStartPosition.CenterScreen;
            ClientSize         = new Size(700, 180);   // ← 改寬

            // 計算上個月的民國年與月份
            var lastMonth = DateTime.Now.AddMonths(-1);
            int rocYear   = lastMonth.Year - 1911;
            int month     = lastMonth.Month;

            // 組出預設的 UNC 路徑
            string networkRoot = @"\\10.21.2.61\藥劑科資料夾\●藥庫\★藥衛材月報表";
            string yearFolder  = $"{rocYear}年";
            string fileName    = $"{rocYear}年{month}月報表.xlsx";
            string defaultPath = Path.Combine(networkRoot, yearFolder, fileName);

            // 年份 Label + NumericUpDown
            Controls.Add(new Label { Text = "年 (民國)：", Location = new Point(20, 20), AutoSize = true });
            nudYear.Minimum   = 1;
            nudYear.Maximum   = 9999;
            nudYear.Value     = rocYear;
            nudYear.Width     = 80;
            nudYear.Location  = new Point(100, 16);
            Controls.Add(nudYear);

            // 月份 Label + NumericUpDown
            Controls.Add(new Label { Text = "月 (1–12)：", Location = new Point(200, 20), AutoSize = true });
            nudMonth.Minimum  = 1;
            nudMonth.Maximum  = 12;
            nudMonth.Value    = month;
            nudMonth.Width    = 50;
            nudMonth.Location = new Point(280, 16);
            Controls.Add(nudMonth);

            // 檔案路徑 Label + TextBox + 瀏覽按鈕
            Controls.Add(new Label { Text = "Excel 檔案：", Location = new Point(20, 60), AutoSize = true });

            txtPath.ReadOnly   = true;
            txtPath.Width      = 600;                                   // ← 改寬
            txtPath.Location   = new Point(20, 80);
            txtPath.Text       = defaultPath;
            txtPath.Anchor     = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right; 
            Controls.Add(txtPath);

            btnBrowse.Text     = "瀏覽...";
            btnBrowse.Width    = 50;
            btnBrowse.Location = new Point(630, 78);                   // ← 改位置
            btnBrowse.Anchor   = AnchorStyles.Top | AnchorStyles.Right;
            btnBrowse.Click   += BtnBrowse_Click;
            Controls.Add(btnBrowse);

            // 確定與取消按鈕
            btnOK.Text         = "確定";
            btnOK.Width        = 75;
            btnOK.Location     = new Point(510, 130);                  // ← 改位置
            btnOK.Anchor       = AnchorStyles.Bottom | AnchorStyles.Right;
            btnOK.Click       += BtnOK_Click;
            Controls.Add(btnOK);

            btnCancel.Text     = "取消";
            btnCancel.Width    = 75;
            btnCancel.Location = new Point(600, 130);                  // ← 改位置
            btnCancel.Anchor   = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCancel.Click   += (s, e) => DialogResult = DialogResult.Cancel;
            Controls.Add(btnCancel);
        }

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Title            = "請選擇 Excel 檔案",
                Filter           = "Excel 檔案 (*.xlsx;*.xls)|*.xlsx;*.xls",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = dlg.FileName;
            }
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPath.Text))
            {
                MessageBox.Show("請先選擇一個 Excel 檔案。", "錯誤",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Year     = (int)nudYear.Value;
            Month    = (int)nudMonth.Value;
            FilePath = txtPath.Text;
            DialogResult = DialogResult.OK;
        }
    }
}
