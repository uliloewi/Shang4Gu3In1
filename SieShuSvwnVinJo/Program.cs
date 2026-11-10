// See https://aka.ms/new-console-template for more information
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using System.Text;

class Program
{

    const string uen2jän4ja5 = @"D:\MyDocument\音韻學\st sk\探索圓脣無介音W\liänmiän\";
    const string excelFile = "上古韻腳2610121046.xlsx";
    const string wordFile= "Sie3Shu5.docx";
    const string filePath = "Sie3Shu5.yaml";
    const string di4i5IPA = "bdɣghklmnŋpʀstx";//所有上古音節中允許的第一個IPA
    static readonly Dictionary<string, string> fen1üin4bu4 = new() //分韻部
    {
            { "歌0", "歌"}, { "歌1","罷"},
            { "月0", "月"}, { "月1","列"},
            { "元0", "元"}, { "元1","仙"},
            { "真0", "真"}, { "真1","陳"},
            { "葉0", "葉"}, { "葉1","蝶"},
            { "談0", "談"}, { "談1","玷"},
    };

    private static Dictionary<string, string> liou3mang4vin4bu4 = new Dictionary<string, string>() {//柳漫韻部
              { "鐸", "ak"}, { "錫", "ɛk"}, { "屋", "ɔk"}, { "職", "ək"}, { "藥", "øk"}, { "覺", "ok"},
              { "陽", "aŋ"}, { "耕", "ɛŋ"}, { "東", "ɔŋ"}, { "蒸", "əŋ"}, { "冬", "oŋ"},
              { "歌", "al"}, { "罷","ɛl"},  { "微", "əl"}, { "脂", "øl"},
              { "月", "at"}, { "列", "ɛt"}, { "物", "ət"}, { "質", "øt"},
              { "元", "an"}, { "仙","ɛn"}, { "文", "ən"},
              { "真", "øn"}, { "陳","øŋ"},
              { "葉", "ap"}, { "蝶","ɛp"}, { "緝", "əp"},
              { "談", "am"}, { "玷","ɛm"}, { "侵", "əm"},
              { "魚", "a"}, { "之", "ə"}, { "支", "ɛ"},
              { "侯", "ɔ"}, { "幽", "o"}, { "宵", "ø"},
        };

    static async System.Threading.Tasks.Task Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        List<string> strings = new List<string>();
        await Du5Chu5Biao3Gä5(excelFile, strings); 

    }

    static async System.Threading.Tasks.Task Du5Chu5Biao3Gä5(string exelFielname, List<string> strings)//讀出表格
    {
        // Pfade zu den Dateien definieren


        try
        {

            Workbook workbook = new Workbook(uen2jän4ja5 + exelFielname);

            int den3er4shr5 = 0; //等二十表即合并一word

            // Durch alle Tabellenblätter der Excel-Datei iterieren
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {               
                den3er4shr5++;
                Worksheet worksheet = workbook.Worksheets[i];

                if (worksheet.Name[0] == worksheet.Name[1])
                {
                   strings.Add($"{worksheet.Name[0]}部韻腳：");
                    Chao1Ie5Biao3(worksheet,  1, worksheet.Name[0].ToString(),  strings);                   
                }
                else if (worksheet.Name[1] == '0')
                {
                   strings.Add($"{fen1üin4bu4[worksheet.Name]}部（{worksheet.Name[0]}甲）韻腳：");
                    Chao1Ie5Biao3(worksheet,  1, worksheet.Name[0].ToString(), strings);
                }
                else if (worksheet.Name[1] == '1')
                {
                   strings.Add($"{fen1üin4bu4[worksheet.Name]}部（{worksheet.Name[0]}乙）韻腳：");
                    Chao1Ie5Biao3(worksheet,  1, fen1üin4bu4[worksheet.Name], strings);
                }
                else if (worksheet.Name[1] == '混')
                {
                   strings.Add( $"{worksheet.Name[0]}之甲乙部通韻：" );
                    //Chao1Ie5Biao3(worksheet, wordApp, 3);
                }
                else
                {
                   strings.Add( $"{worksheet.Name}二部通韻：" );
                    //Chao1Ie5Biao3(worksheet, wordApp, 4);
                }
            }
            
             await File.WriteAllLinesAsync(uen2jän4ja5 + "800" + filePath, strings);
            Console.WriteLine("Die Daten wurden erfolgreich in das Word-Dokument übertragen!");
        }

        catch (Exception ex)
        {
            Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}");
        }
    }

    private static void Chao1Ie5Biao3(Worksheet worksheet,  int biao3xin2, string üin4bu4, List<string> strings)//抄頁表 表型biao3xin2=1如：魚魚，biao3xin2=2如：月0，biao3xin2=3如：月混，biao3xin2=4如：魚鐸
    {
        int hang2 = 0;
        while (!(worksheet.Cells[hang2, 1].StringValue == "" && worksheet.Cells[hang2+1, 1].StringValue == "" && worksheet.Cells[hang2 + 2, 1].StringValue == ""))
        {
            var stringA = worksheet.Cells[hang2, 0].StringValue;
            var piän1min2 = stringA.Substring(stringA.IndexOf("-") + 1).Replace("-", "・");//篇名
            if (!String.IsNullOrWhiteSpace(piän1min2))
            {
                string shr1hang2 = "《" + piän1min2 + "》： ";//詩行
                
                Console.WriteLine(piän1min2);
                if (piän1min2.Contains("汝墳") )
                {

                }
                switch (biao3xin2)//表型
                {
                    case 1://如：魚魚
                        Süän3In1(worksheet, liou3mang4vin4bu4[üin4bu4], hang2, strings, shr1hang2);   
                        //wordApp.Selection.TypeText(worksheet.Cells[hang2, 1].StringValue + "、" + worksheet.Cells[hang2, 2].StringValue + "、" + worksheet.Cells[hang2, 3].StringValue + "、" + worksheet.Cells[hang2, 4].StringValue + "、" + worksheet.Cells[hang2, 5].StringValue + "、" + worksheet.Cells[hang2, 6].StringValue + "\n");
                        break;
                    case 2://如：月0
                        //wordApp.Selection.TypeText(worksheet.Cells[hang2, 1].StringValue + "、" + worksheet.Cells[hang2, 2].StringValue + "、" + worksheet.Cells[hang2, 3].StringValue + "、" + worksheet.Cells[hang2, 4].StringValue + "、" + worksheet.Cells[hang2, 5].StringValue + "、" + worksheet.Cells[hang2, 6].StringValue + "\n");
                        break;
                    case 3://如：月混
                        //wordApp.Selection.TypeText(worksheet.Cells[hang2, 1].StringValue + "、" + worksheet.Cells[hang2, 2].StringValue + "、" + worksheet.Cells[hang2, 3].StringValue + "、" + worksheet.Cells[hang2, 4].StringValue + "、" + worksheet.Cells[hang2, 5].StringValue + "、" + worksheet.Cells[hang2, 6].StringValue + "\n");
                        break;
                    case 4://如：魚鐸
                        //wordApp.Selection.TypeText(worksheet.Cells[hang2, 1].StringValue + "、" + worksheet.Cells[hang2, 2].StringValue + "、" + worksheet.Cells[hang2, 3].StringValue + "、" + worksheet.Cells[hang2, 4].StringValue + "、" + worksheet.Cells[hang2, 5].StringValue + "、" + worksheet.Cells[hang2, 6].StringValue + "\n");
                        break;
                    default:
                        break;
                }
            }
            hang2++;
        }
    }

    private static void Süän3In1(Worksheet worksheet,  string üin4bu4,int hang2, List<string> strings, string shr1hang2)//選音
    {
        try
        {
            Dictionary<int, string> üin4jo5zy4 = new Dictionary<int, string>();//韻腳字，如1.有，2.祀
            Dictionary<int, List<string>> do1go4in1 = new Dictionary<int, List<string>>();//多個音如1.ɣʷəɣ	、ɣʷəh，2.sɣəɣ
            Dictionary<int, List<string>> zy4ho2in1 = new Dictionary<int, List<string>>();//字和音
            int col = 1;
            int key = 0;
            while (!(worksheet.Cells[hang2, col].StringValue == "" && worksheet.Cells[hang2, col + 1].StringValue == "" && worksheet.Cells[hang2, col + 2].StringValue == ""))
            {
                var stringA = worksheet.Cells[hang2, 0].StringValue;
                var piän1min2 = stringA.Substring(stringA.IndexOf("-") + 1).Replace("-", "・");//篇名
                if (!String.IsNullOrWhiteSpace(piän1min2))
                {
                    stringA = worksheet.Cells[hang2, col].StringValue;
                    col++;
                    if (String.IsNullOrEmpty(stringA))
                        continue;
                    if (!di4i5IPA.Contains(stringA[0]))
                    {
                        if (stringA.Length <= 2 || stringA.Contains("（"))//是漢字
                        {
                            key++;
                            üin4jo5zy4.Add(key, stringA);
                        }
                        else//是注
                        {
                           strings.Add( stringA + "   注！！！  \n" );
                            break;
                        }
                    }
                    else//是IPA，做value
                    {
                        if (do1go4in1.ContainsKey(key))
                        {
                            do1go4in1[key].Add(stringA);
                        }
                        else
                        {
                            do1go4in1.Add(key, new List<string> { stringA });
                        }

                    }
                }
            }
            An4Diao4Süän3In1(üin4jo5zy4, do1go4in1, üin4bu4, "h");
            An4Diao4Süän3In1(üin4jo5zy4, do1go4in1, üin4bu4, "ɣ");
            An4Diao4Süän3In1(üin4jo5zy4, do1go4in1, üin4bu4, "");
            for (int i = 1;i<=üin4jo5zy4.Count;i++)
            {
                shr1hang2+= üin4jo5zy4[i];
                Console.WriteLine(üin4jo5zy4[i]);
            }
           strings.Add(shr1hang2);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}");
        }
    }


    private static void An4Diao4Süän3In1(Dictionary<int, string> üin4jo5zy4, Dictionary<int, List<string>> do1go4in1, string üin4bu4, string shen1diao4uei3)
    {
        try { 
            if (do1go4in1.All(x => x.Value.Any(y => y.EndsWith(üin4bu4 + shen1diao4uei3))))
            {//每個字都有去聲韻尾
                for (int i = 1; i <= do1go4in1.Count; i++)
                {
                    string zy4in1 = "";
                    if (do1go4in1[i].Count > 1)
                    {
                        foreach (var x in do1go4in1[i])
                        {
                            if (x.EndsWith(üin4bu4 + shen1diao4uei3))
                                zy4in1 += x + "，";
                        }
                    }
                    else
                    {
                        zy4in1 += do1go4in1[i][0] + "，";
                    }

                    üin4jo5zy4[i]+=zy4in1;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}");
        }
    }
}

