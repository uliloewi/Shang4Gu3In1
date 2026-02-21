using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using static System.Net.Mime.MediaTypeNames;
namespace Shang4Gu3In1
{
    /// <summary>
    ///    檢查連綿詞擬音阿合理
    /// </summary>
    public class Die5Yuin4
    {
        const string lu5vwn2in1 = "aɛɔəøo";//六元音
        const string ci5vin4uei3 = "ŋkmpntl";//七韻尾
        List<string> shen1mu3 = new List<string>() {
            "ʀk", "ʀg", "ʀx",
            "skʀ", "skʰʀ", "sgʀ", "sxʀ",
            "stʀ", "stʰʀ", "sdʀ",
            "skʰ", "sk", "sg",
            "stʰ", "st", "sd",
            "sx", "sŋ",
            "tʀ", "tʰʀ", "dʀ", "nʀ",
            "sʀ",
            "kl", "kʰl", "gl", "ŋl",
            "xl", "xn", "hl",
             "ɣl", "l",
            "tʰ", "t", "d",
            "n",
            "ʀ",
            "s",
            "pʰ", "p", "b",
            "m",
            "h",
            "x",
            "ɣ",
            "kʰ", "k",  "g",
            "ŋ"};

        const string uen2jän4ja5 = @"C:\Users\xggg\Downloads\SieShu\";
        public static void Main1(string[] args, string uen2jän4ja5)
        {

            Console.OutputEncoding = Encoding.UTF8;
            Gai3Zhong1Gu3Pin1In1(uen2jän4ja5);
            Workbook wk = new Workbook(uen2jän4ja5 + "a.xlsx");
            Worksheet ws = wk.Worksheets[0];
            /*
            foreach (string line in File.ReadLines(uen2jän4ja5 + "dazydiän.txt"))
            {
                bool tong2üän2 = false;
                int idxBlank = line.IndexOf(' ');
                if (idxBlank > 0)
                {
                    string zy = line.Substring(0, idxBlank);
                    if (zy.Contains('('))
                    {
                        tong2üän2 |= true;
                        zy = zy.Replace("(", "").Replace(")", "");
                    }
                    zy += GetTongJia(line, "，同“", ref tong2üän2);
                    zy += GetTongJia(line, "，通“", ref tong2üän2);
                    if (tong2üän2)
                    {
                        NeedNotice(zy, ws);
                    }
                }
            }*/

            List<string> lm = new List<string>();
            foreach (string line in File.ReadLines(uen2jän4ja5 + "連綿詞.txt"))
            {
                var x = line.Split("/");
                foreach (string s in x)
                {
                    if (!lm.Contains(s))
                    {
                        if (!Regex.IsMatch(s, "[a-zA-Z]"))
                        {
                            NeedNotice(s, ws, true);
                        }
                        lm.Add(s);
                    }
                }
            }


            static string GetTongJia(string line, string splitter, ref bool tong2üän2)
            {
                string res = "";
                var tong = line.Split(splitter);//如: "，同“"
                if (tong.Length > 1)
                {
                    tong2üän2 = true;
                    for (int i = 1; i < tong.Length; i++)
                    {
                        if (tong[i].IndexOf("”") > 0)
                            res += tong[i].Substring(0, tong[i].IndexOf("”"));
                    }
                }
                return res;
            }

            static void NeedNotice(string zy, Worksheet ws, bool chu4li3liwn2miwn2 = false)
            {
                if (zy[0].ToString() == zy[1].ToString()) return;
                bool res = false;
                List<string> listZy = new List<string>();

                for (int j = 0; j < zy.Length; j++)
                {
                    string danzy = zy[j].ToString();
                    if (j + 1 < zy.Length &&
                        Encoding.UTF8.GetBytes(zy[j].ToString())[0] == Encoding.UTF8.GetBytes(zy[j + 1].ToString())[0] &&
                        Encoding.UTF8.GetBytes(zy[j].ToString())[1] == Encoding.UTF8.GetBytes(zy[j + 1].ToString())[1] &&
                        Encoding.UTF8.GetBytes(zy[j].ToString())[2] == Encoding.UTF8.GetBytes(zy[j + 1].ToString())[2])
                    {
                        danzy = zy[j].ToString() + zy[j + 1].ToString();
                        j++;
                    }
                    listZy.Add(danzy);
                }
                if (listZy.Count > 0 && listZy[0] == listZy[1])
                    return;
                Dictionary<string, List<string>> map = ZyDict(listZy);
                Dictionary<string, List<string>> sheng1 = ZyDict(listZy);
                Dictionary<string, List<string>> in1zie5 = ZyDict(listZy);
                Dictionary<string, List<string>> cie5in1 = ZyDict(listZy);

                for (int row = 3; row <= ws.Cells.MaxDataRow; row++)
                {//處理了一字所有古音
                    if (ws.Cells["P" + row.ToString()].Value == null || ws.Cells["G" + row.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                        continue;
                    foreach (var a in listZy)
                    {
                        if (ws.Cells["P" + row.ToString()].Value.ToString().Contains(a))//"P"列是同聲旁同音字
                        {
                            map[a].Add(ws.Cells["E" + row.ToString()].Value?.ToString());
                            sheng1[a].Add(ws.Cells["H" + row.ToString()].Value?.ToString());
                            in1zie5[a].Add(ws.Cells["G" + row.ToString()].Value?.ToString());
                            cie5in1[a].Add(ws.Cells["N" + row.ToString()].Value?.ToString());
                        }
                    }
                }
                if ((chu4li3liwn2miwn2 && !Shr4Die5Vin4(in1zie5) && !Shr4Shuang1Shen1(in1zie5)) ||
                    (!chu4li3liwn2miwn2 && !AllListsShareCommonElement(in1zie5, true)))
                {
                    for (int i = 0; i < listZy.Count; i++)
                    {

                        Console.Write(listZy[i]);
                        Console.Write(string.Join(";", in1zie5[listZy[i]]));
                        if (i == listZy.Count - 1)
                            Console.WriteLine();
                    }
                    string ciein = "";
                    for (int i = 0; i < listZy.Count; i++)
                    {
                        ciein += string.Join(";", cie5in1[listZy[i]]) + " && ";
                    }
                    Console.WriteLine(ciein);
                }
            }

            static Dictionary<string, List<string>> ZyDict(List<string> listZy)
            {
                Dictionary<string, List<string>> map = new Dictionary<string, List<string>>();
                if (listZy.Count > 1 && listZy[0] != listZy[1])
                {
                    foreach (string s in listZy)
                    {
                        map.Add(s, new List<string>());
                    }
                }
                return map;
            }

            static bool AllListsShareCommonElement(Dictionary<string, List<string>> lists, bool ignoreEmptyList = false)
            {
                if (lists == null || lists.Count == 0) return false;
                Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
                if (ignoreEmptyList)
                {
                    foreach (var l in lists.Where(x => x.Value.Count > 0))
                        dic.Add(l.Key, l.Value);
                }
                else
                {
                    dic = lists;
                }

                // Wir starten mit der ersten Liste als Basis für die Schnittmenge
                if (dic.Count == 0)//如：焙𤊷不在字表中
                    return true;

                return Iou3Gong4Tong2Vwn2Su4(dic).Any();
            }

            static string GetWord(string input, string a, string b)
            {
                string result = "";
                int end = input.IndexOf(b); // Find the first colon after that space
                if (end > 1)
                {
                    string sub = input.Substring(0, end);
                    int start = sub.LastIndexOf(a) + 1;

                    if (start > 0 && end > start)
                    {
                        result = input.Substring(start, end - start);
                    }
                }
                return result;
            }

            static bool Shr4Die5Vin4(Dictionary<string, List<string>> lists)//是叠韻
            {
                Dictionary<string, List<string>> uei3zy4mu3 = Duang1String(lists, 3);
                if (Iou3Gong4Tong2Vwn2Su4(uei3zy4mu3).Any())
                    return true;
                else
                {
                    uei3zy4mu3 = Duang1String(lists, 2);
                    if (lu5vwn2in1.Any(v => Iou3Gong4Tong2Vwn2Su4(uei3zy4mu3).Any(x => x.Contains(v))))
                        return true;
                    else
                    {
                        uei3zy4mu3 = Duang1String(lists, 1);
                        if (lu5vwn2in1.Any(v => Iou3Gong4Tong2Vwn2Su4(uei3zy4mu3).Any(x => x.Contains(v))))
                            return true;
                    }
                }
                return false;
            }

            static Dictionary<string, List<string>> Duang1String(Dictionary<string, List<string>> dic, int length, bool zhao3kw1tou2 = false)
            {
                Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
                foreach (var l in dic)
                {
                    res[l.Key] = new List<string>();
                    foreach (var inzie in l.Value)
                    {
                        if (zhao3kw1tou2)
                            res[l.Key].Add(inzie.Substring(0, length));
                        else if (inzie.Length >= length)
                            res[l.Key].Add(inzie.Substring(inzie.Length - length));
                    }
                }
                return res;
            }


            static IEnumerable<string> Iou3Gong4Tong2Vwn2Su4(Dictionary<string, List<string>> dic)//有共同元素
            {
                IEnumerable<string> intersection = dic.First().Value;

                // Wir vergleichen die Basis nacheinander mit jeder weiteren Liste
                for (int i = 1; i < dic.Count; i++)
                {
                    intersection = intersection.Intersect(dic.ElementAt(i).Value);
                }

                // Wenn am Ende noch Elemente übrig sind, gibt es ein gemeinsames Element in ALLEN Listen
                return intersection;
            }

            static bool Shr4Shuang1Shen1(Dictionary<string, List<string>> lists)//是雙聲
            {
                return Iou3Gong4Tong2Vwn2Su4(Cv3Sheng1Mu3(lists)).Any();
            }

            static string GetOnset(string in1zie5, bool checkRK = false)//sgat->sg
            {
                List<string> vin4mu3lie5bao3 = new List<string>();//所有上古韻母
                foreach (var item in lu5vwn2in1)
                {
                    vin4mu3lie5bao3.Add("ˤ" + item);
                    vin4mu3lie5bao3.Add(item.ToString());
                }
                string res = "";
                foreach (var v in vin4mu3lie5bao3)
                {
                    if (in1zie5.IndexOf(v) > 0)
                    {
                        res = in1zie5.Substring(0, in1zie5.IndexOf(v));
                    }
                }
                return checkRK ? res.Replace("tʀ", "rk").Replace("dʀ", "rg").Replace("tʰʀ", "rx").Replace("nʀ", "rŋ") : res;
            }

            static Dictionary<string, List<string>> Cv3Sheng1Mu3(Dictionary<string, List<string>> dic)
            {
                Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
                foreach (var l in dic)
                {
                    res[l.Key] = new List<string>();
                    foreach (var inzie in l.Value)
                    {
                        res[l.Key].Add(GetOnset(inzie).Replace("ʷ", "").Replace("ˤ", ""));
                    }
                }
                return res;
            }

            ///把大字典中的古韻拼音改成廣通拼音
            static void Gai3Zhong1Gu3Pin1In1(string uen2jän4ja5)//改中古拼音
            {
                Workbook wk = new Workbook(uen2jän4ja5 + "Guangyun_Langjin_pulish_Alphabetic.2.0.xlsx");
                var gu3vin4 = Fang3Cie5Duei4Pin1In1(wk, "H", "K");
                wk = new Workbook(uen2jän4ja5 + "Guangyun_Langjin_Zhonggu.3.0.xlsx");
                var guang3tong1 = Fang3Cie5Duei4Pin1In1(wk, "I", "K");
                Dictionary<string, string> dic = new Dictionary<string, string>();
                foreach (var l in gu3vin4)
                {
                    if (l.Key != "十合" && l.Key != "辝纂")
                        dic[l.Value] = guang3tong1[l.Key.Replace("愽", "博").Replace("别", "別").Replace("没", "沒")
                            .Replace("卧", "臥").Replace("鑒", "鑑").Replace("居帋", "居氏").Replace("叉万", "初万").Replace("博耗", "博秏")];//key:mox,value:mó
                }
                /*
                using (WordprocessingDocument doc = WordprocessingDocument.Open(uen2jän4ja5 + "m.docx", true))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    // Wir suchen nach allen "Text"-Elementen im Dokument
                    foreach (var textElement in body.Descendants<Text>())
                    {
                        foreach (var d in dic.OrderByDescending(x => x.Key.Length))
                        {
                            // Wichtig: Wir ändern NUR die .Text Eigenschaft.
                            // Die übergeordneten RunProperties (Style, Size, etc.) bleiben unberührt.
                            textElement.Text = textElement.Text.Replace(d.Key, d.Value);
                        }
                    }

                    doc.MainDocumentPart.Document.Save();
                }*/

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(uen2jän4ja5 + "dazydiän.docx");
                for (int i = 1; i < wordDoc.Paragraphs.Count; i++)
                {
                    var txt = wordDoc.Paragraphs[i].Range.Text;
                    int idx = txt.IndexOf("中古音：");
                    if (idx > 0)
                    {
                        var in1zie5 = txt.Substring(idx + 4).Replace("\r", "");
                        var in1 = in1zie5.Split("；");
                        foreach (var cie5fa1in1 in in1)
                        {
                            var k = cie5fa1in1.Split("切");//莫六“切”miuk
                            if (k.Length > 1)
                            {
                                if (dic.Keys.Contains(k[1]))
                                    txt = txt.Replace(k[1], dic[k[1]]);
                                else
                                    Console.WriteLine(k[1] + " no mapping");
                            }
                        }
                        if (txt.EndsWith(";"))
                            txt = txt.Substring(0, txt.Length - 1);
                        wordDoc.Paragraphs[i].Range.Text = txt;
                        wordDoc.Paragraphs[i].Range.Font.Size = 10;
                        var fen1zy4tou2 = txt.Split(" ");
                        var firstPart = wordDoc.Range(wordDoc.Paragraphs[i].Range.Start, wordDoc.Paragraphs[i].Range.Start + fen1zy4tou2[0].Length);
                        firstPart.Font.Size = 21;
                    }
                }
                wordDoc.Save();
            }

            static Dictionary<string, string> Fang3Cie5Duei4Pin1In1(Workbook wk, string colCie, string colPinIn)
            {
                Worksheet ws = wk.Worksheets[0];
                Dictionary<string, string> vs = new Dictionary<string, string>();
                for (int row = 2; row <= ws.Cells.MaxDataRow; row++)
                {
                    if (ws.Cells[colCie + row.ToString()].Value != null && ws.Cells[colPinIn + row.ToString()].Value != null)
                        vs.Add(ws.Cells[colCie + row.ToString()].Value.ToString().Trim(), ws.Cells[colPinIn + row.ToString()].Value.ToString().Trim());
                }
                return vs;
            }
        }
    }
}