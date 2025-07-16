// 先在專案中安裝 NuGet 套件：DocumentFormat.OpenXml
// PM> Install-Package DocumentFormat.OpenXml

using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MyApp.Utils
{
    public static class ReportWriter
    {
        /// <summary>
        /// 傳入 年、月、以及 34 個欄位值，
        /// 檢查長度、依序填入模板中所有 '@'，
        /// 並將結果存成 Word (docx) 檔到使用者桌面。
        /// </summary>
        public static void WriteReport(int year, int month, List<string> values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            if (values.Count != 34)
                throw new ArgumentException("values 必須包含 34 項", nameof(values));

            // 原始模板，所有需要被替換的欄位都用「@」表示
            var template = @"主旨：陳本院@年@月份藥品及衛材耗用金額總報表，如說明，請鑒核。

說明：
藥品門診耗用成本@元，住院耗用成本@元，診所耗用成本@元，花蓮榮家耗用成本@元，合計@元。

衛材門診耗用成本@元，住院耗用成本@元，診所耗用成本@元，花蓮榮家耗用成本@元，合計@元。

門診醫療收入為@元。門診藥品耗用成本@元，佔門診收入@，門診衛材耗用成本@元，佔門診收入@。

住院醫療收入@元。住院藥品耗用成本@元，佔住院收入@，住院衛材耗用成本@元，佔住院收入@。

藥品及衛材耗用量總金額@元，門診住院醫療總收入@元，藥衛材耗用成本佔全院總收入@。

本院門住院藥品耗用成本@元，佔本院門住院醫療收入@元，比率@；本院門住院診衛材耗用成本@元，佔本院門住院醫療收入@，本院門住院藥衛材耗用成本@元，佔本院門住院醫療收入@。

藥品庫存比@；衛材庫存比@；藥衛材庫存比@，藥衛材週轉率@。

擬辦：
奉核後文存備查。";

            // 1) 把年、月、再加上 34 個 values，放進一個替換清單
            var tokens = new List<string> { year.ToString(), month.ToString() };
            tokens.AddRange(values);

            // 2) 依序找到第一個 '@'，拿 tokens[0] 取代，再找下一個 '@' 用 tokens[1]，以此類推
            var filled = FillTemplate(template, tokens);

            // 3) 存成 Word (docx) 到桌面
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var fileName = $"報表_{year}_{month}.docx";
            var outputPath = Path.Combine(desktop, fileName);
            CreateWordDocument(outputPath, filled);
        }

        // 取代第一個 '@' → tokens[0]，再替換下一個 → tokens[1]，依序下去
        private static string FillTemplate(string template, List<string> tokens)
        {
            var result = template;
            foreach (var t in tokens)
            {
                var idx = result.IndexOf('@');
                if (idx < 0) break;
                result = result.Substring(0, idx) + t + result.Substring(idx + 1);
            }
            return result;
        }

        // 用 Open XML SDK 建立一個簡單的 .docx，並逐行當作獨立段落
        private static void CreateWordDocument(string filePath, string content)
        {
            if (File.Exists(filePath))
                File.Delete(filePath);

            using var wordDoc = WordprocessingDocument.Create(
                filePath,
                WordprocessingDocumentType.Document);

            // 1. 新建 Document 實例
            var document = new Document();

            // 2. 新建 Body，再附加到 Document
            var body = new Body();
            document.Append(body);

            // 3. 把我們的 Document 套給 mainPart
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = document;

            // 4. 產生段落並加入 body（此時 body 絕對不為 null）
            var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                var p = new Paragraph(new Run(new Text(line) { Space = SpaceProcessingModeValues.Preserve }));
                body.Append(p);
            }

            mainPart.Document.Save();
        }
    }
}
