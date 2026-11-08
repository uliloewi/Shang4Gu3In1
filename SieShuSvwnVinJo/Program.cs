// See https://aka.ms/new-console-template for more information
using Aspose.Cells;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

class Program
{

    const string uen2jän4ja5 = @"D:\MyDocument\音韻學\st sk\探索圓脣無介音W\liänmiän\";
    const string excelFile = "上古韻腳.xlsx";
    const string wordFile = "800Sie3Shu5.docx";
    static readonly Dictionary<string, string> fen1üin4bu4 = new() //分韻部
    {
            { "歌0", "歌"}, { "歌1","罷"},  
            { "月0", "月"}, { "月1","列"}, 
            { "元0", "元"}, { "元1","仙"},
            { "真0", "真"}, { "真1","陳"},
            { "葉0", "葉"}, { "葉1","蝶"}, 
            { "談0", "談"}, { "談1","玷"}, 
    };

    static async System.Threading.Tasks.Task Main(string[] args)
    {
        Application wordApp = new Application();
        Document doc = wordApp.Documents.Add();
        Du5Chu5Biao3Gä5(excelFile, wordFile, wordApp, doc);
    }

    private static void Du5Chu5Biao3Gä5(string exelFielname, string wordFielname, Application wordApp, Document doc)//讀出表格
    {
        // Pfade zu den Dateien definieren


        try
        {
            // 1. Excel-Datei mit Aspose.Cells einlesen
            Workbook workbook = new Workbook(uen2jän4ja5 + exelFielname);


            // Durch alle Tabellenblätter der Excel-Datei iterieren
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet worksheet = workbook.Worksheets[i];

                if (worksheet.Name[0] == worksheet.Name[1])
                {
                    wordApp.Selection.TypeText($"{worksheet.Name[0]}部韻腳：\n");
                }
                else if (worksheet.Name[1] == '0' )
                {
                    wordApp.Selection.TypeText($"{fen1üin4bu4[worksheet.Name]}部（{worksheet.Name[0]}甲）韻腳：\n");
                }
                else if (worksheet.Name[1] == '1')
                {
                    wordApp.Selection.TypeText($"{fen1üin4bu4[worksheet.Name]}部（{worksheet.Name[0]}乙）韻腳：\n");
                }
                else if (worksheet.Name[1] == '混')
                {
                    wordApp.Selection.TypeText($"{worksheet.Name[0]}之甲乙部通韻：\n");
                }
                else
                {
                    wordApp.Selection.TypeText($"{worksheet.Name}二部通韻：\n");
                }
                int hang2 = 0;
                while (worksheet.Cells[hang2, 1].StringValue != "")
                {
                    var stringA = worksheet.Cells[hang2, 0].StringValue;
                    var piän1min2 = stringA.Substring(stringA.IndexOf("-")+1).Replace("-", "・");//篇名
                    wordApp.Selection.TypeText("《"+piän1min2+ "》：\n");
                    hang2++;
                }
                //// Werte aus den Zellen B1 und C1 auslesen
                //// (Zeilen- und Spaltenindizes sind 0-basiert: B1 = R0C1, C1 = R0C2)
                //string cellB1Text = worksheet.Cells[0, 1].StringValue;
                //string cellC1Text = worksheet.Cells[0, 2].StringValue;

                //// Text in das Word-Dokument schreiben (Ein Text pro Zeile)
                //wordApp.Selection.TypeText($"Zelle B1: {cellB1Text}");
                //wordApp.Selection.TypeText($"Zelle C1: {cellC1Text}");

                // Wenn es nicht das letzte Blatt ist, einen Seitenumbruch einfügen
                if (i < workbook.Worksheets.Count - 1)
                {
                    object breakPage = Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                    wordApp.Selection.InsertBreak(ref breakPage);
                }
            }

            // 3. Word-Dokument speichern
            doc.SaveAs2(uen2jän4ja5 + wordFielname);
            Console.WriteLine("Die Daten wurden erfolgreich in das Word-Dokument übertragen!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}");
        }
        finally
        {
            // 3. Wichtig: COM-Objekte sauber schließen und freigeben
            if (doc != null)
            {
                doc.Close(false);
                Marshal.ReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}

