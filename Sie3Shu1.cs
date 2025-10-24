using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Drawing;

namespace Shang4Gu3In1
{
    /* 寫書 */
    public static class Sie3Shu1
    {
        private static Dictionary<int, double> kuang1du4 = new Dictionary<int, double>() {
            { 1, 15},//直聲旁
            { 2, 15},//韻部
            { 3, 30},//演變
            { 4, 35},//曹魯音
            { 5, 15},//聲
            { 6, 15},//等
            { 7, 15},//呼
            { 8, 25},//韻
            { 9, 15},//調
            { 10, 30},//切
            { 11, 35},//中古
            { 12, 35},//南京
            { 13, 60},//字
            { 14, 60}//註
        };
        /// AI created
        /// Kopiert einen Zellbereich aus einer Excel-Datei in ein Word-Dokument als Tabelle.
        /// startRow und startCol sind 1-basiert (wie in Ihrem Aufruf).
        /// </summary>
        public static void CopyRangeToWord(string excelFilePath, int sheetIdx, int startRow, int startCol, int rowCount, int colCount, string outputWordFilePath)
        {
            if (string.IsNullOrEmpty(excelFilePath)) throw new ArgumentNullException(nameof(excelFilePath));
            if (string.IsNullOrEmpty(outputWordFilePath)) throw new ArgumentNullException(nameof(outputWordFilePath));
            // Laden der Excel-Datei
            var workbook = new Workbook(excelFilePath);
            // Finde Arbeitsblatt per Name, fallback auf erstes Blatt
            Worksheet worksheet = workbook.Worksheets[sheetIdx];

            // Indizes in Aspose.Cells sind 0-basiert
            int r0 = Math.Max(0, startRow - 1);
            int c0 = Math.Max(0, startCol - 1);

            // Begrenze Bereich auf vorhandene Zeilen/Spalten
            int maxRow = Math.Min(r0 + Math.Max(0, rowCount) - 1, worksheet.Cells.MaxDataRow);
            int maxCol = Math.Min(c0 + Math.Max(0, colCount) - 1, worksheet.Cells.MaxDataColumn);

            // Neues Word-Dokument
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Erstelle Tabelle
            var table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            for (int r = r0; r <= maxRow; r++)
            {
                var wordRow = new Aspose.Words.Tables.Row(doc);
                for (int c = c0; c <= maxCol; c++)
                {
                    var excelCell = worksheet.Cells[r, c];
                    string text = excelCell?.StringValue ?? string.Empty;

                    var wordCell = new Aspose.Words.Tables.Cell(doc);
                    // Paragraph + Run für Text und Schriftformatierung
                    var para = new Paragraph(doc);
                    var run = new Run(doc, text);

                    // Übernehme einfache Font-Eigenschaften, wenn vorhanden
                    try
                    {
                        var style = excelCell.GetStyle();
                        if (style != null)
                        {
                            var f = style.Font;
                            if (f != null)
                            {
                                run.Font.Name = f.Name ?? run.Font.Name;
                                if (f.Size > 0) run.Font.Size = f.Size;
                                run.Font.Bold = f.IsBold;
                                // Aspose.Cells.Font.Color ist System.Drawing.Color
                                if (f.Color != Color.Empty)
                                {
                                    run.Font.Color = f.Color;
                                }
                            }
                            // Hintergrundfarbe der Zelle (einfach als Shading)
                            if (style.ForegroundColor != Color.Empty)
                            {
                                wordCell.CellFormat.Shading.BackgroundPatternColor = style.ForegroundColor;
                            }
                        }
                    }
                    catch
                    {
                        // Styles sind optional — bei Fehlern weiter ohne Style
                    }

                    para.AppendChild(run);
                    wordCell.AppendChild(para);
                    wordRow.AppendChild(wordCell);
                }
                table.AppendChild(wordRow);
            }

            // Option: einfache Rahmen hinzufügen
            foreach (Aspose.Words.Tables.Row wr in table.Rows)
            {
                foreach (Aspose.Words.Tables.Cell wc in wr.Cells)
                {
                    wc.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                }
            }

            // Speichern als .docx
            doc.Save(outputWordFilePath, Aspose.Words.SaveFormat.Docx);
        }

        public static void ExcelToWordTable(string excelFilePath, int sheetIdx, int startRow, int startCol, int rowCount, int colCount, string outputWordFilePath)
        {
            if (string.IsNullOrEmpty(excelFilePath)) throw new ArgumentNullException(nameof(excelFilePath));
            if (string.IsNullOrEmpty(outputWordFilePath)) throw new ArgumentNullException(nameof(outputWordFilePath));
            // Laden der Excel-Datei
            var workbook = new Workbook(excelFilePath);
            // Finde Arbeitsblatt per Name, fallback auf erstes Blatt
            Worksheet worksheet = workbook.Worksheets[sheetIdx];

            // Indizes in Aspose.Cells sind 0-basiert
            int r0 = Math.Max(0, startRow - 1);
            int c0 = Math.Max(0, startCol - 1);

            // Begrenze Bereich auf vorhandene Zeilen/Spalten
            int maxRow = Math.Min(r0 + Math.Max(0, rowCount) - 1, worksheet.Cells.MaxDataRow);
            int maxCol = Math.Min(c0 + Math.Max(0, colCount) - 1, worksheet.Cells.MaxDataColumn);

            // Neues Word-Dokument
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Erstelle Tabelle
            var table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            for (int r = r0; r <= maxRow; r++)
            {
                int i = 0;
                var wordRow = new Aspose.Words.Tables.Row(doc);
                for (int c = c0; c <= maxCol; c++)
                {
                    i++;
                    var excelCell = worksheet.Cells[r, c];
                    string text = excelCell?.StringValue ?? string.Empty;

                    var wordCell = new Aspose.Words.Tables.Cell(doc);
                    // Paragraph + Run für Text und Schriftformatierung
                    var para = new Paragraph(doc);
                    var characterColors = GetCharacterColors(excelCell);
                    string hong2zy4 = characterColors.Where(c => c.Color.Name == "ffff0000").Select(characterColors => characterColors.Character).Aggregate("", (current, ch) => current + ch);

                    var run = new Run(doc,String.IsNullOrEmpty(hong2zy4)?text:text.Replace(hong2zy4,""));
                    // Übernehme einfache Font-Eigenschaften, wenn vorhanden
                    try
                    {
                        var style = excelCell.GetStyle();

                        if (style != null)
                        {
                            var f = style.Font;
                            if (f != null)
                            {
                                if (text.StartsWith("匹"))
                                {
                                    
                                }
                                if (hong2zy4?.Length > 0)
                                {
                                    var hong2run = new Run(doc, hong2zy4);
                                    hong2run.Font.Color = Color.Red;
                                    hong2run.Font.Name = "SimSun-ExtB";
                                    para.AppendChild(hong2run);
                                }
                                run.Font.Name = f.Name ?? run.Font.Name;
                                if (i >= 13)
                                    run.Font.Name = "SimSun-ExtB";
                                run.Font.Size = 9;
                                //run.Font.Bold = f.Color != Color.Empty && f.Color.Name == "ffff0000";
                                run.Font.Bold = f.IsBold;
                                // Aspose.Cells.Font.Color ist System.Drawing.Color
                                
                                if (f.Color != Color.Empty)
                                {
                                    run.Font.Color = f.Color;
                                }
                            }
                            // Hintergrundfarbe der Zelle (einfach als Shading)
                            if (style.ForegroundColor != Color.Empty)
                            {
                                wordCell.CellFormat.Shading.BackgroundPatternColor = style.ForegroundColor;
                            }
                        }
                    }
                    catch
                    {
                        // Styles sind optional — bei Fehlern weiter ohne Style
                    }

                    para.AppendChild(run);
                    wordCell.AppendChild(para);
                    wordRow.AppendChild(wordCell);
                }
                table.AppendChild(wordRow);
            }
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Option: einfache Rahmen hinzufügen
            
            foreach (Aspose.Words.Tables.Row wr in table.Rows)
            {
                int i = 0;
                foreach (Aspose.Words.Tables.Cell wc in wr.Cells)
                {
                    i++;
                    wc.CellFormat.PreferredWidth = PreferredWidth.FromPoints(kuang1du4[i]);
                    wc.CellFormat.Width = kuang1du4[i];
                    wc.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
                    wc.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                }
            }

            // Speichern als .docx
            doc.Save(outputWordFilePath, Aspose.Words.SaveFormat.Docx);
        }

        public static List<(char Character, Color Color)> GetCharacterColors(Aspose.Cells.Cell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            string text = cell.StringValue ?? (cell.Value?.ToString() ?? string.Empty);
            var result = new List<(char, Color)>();

            // Basis-Font aus Zell-Style
            Color defaultColor = Color.Black;
            try
            {
                var style = cell.GetStyle();
                if (style?.Font != null && style.Font.Color != Color.Empty)
                    defaultColor = style.Font.Color;
            }
            catch { /* ignore */ }

            for (int i = 0; i < text.Length; i++)
            {
                try
                {
                    // GetCharacters verwendet UTF-16 index und Länge in chars
                    var charFont = cell.Characters(i, 1);
                    Color c = (charFont != null && charFont.Font.Color != Color.Empty) ? charFont.Font.Color : defaultColor;
                    result.Add((text[i], c));
                }
                catch
                {
                    result.Add((text[i], defaultColor));
                }
            }

            return result;
        }
    }
}
