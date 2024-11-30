using Shang4Gu3In1;
using Aspose.Cells;
using System.Drawing;
using System.Runtime.CompilerServices;

namespace zhongguliin
{
    class Program
    {
        private static Dictionary<string, string[]> vin14bu4 = new Dictionary<string, string[]>() {
           /* { "魚", ["a"]}, { "鐸", ["ak"]}, { "陽", ["aŋ"]},
            { "之", ["ə"]}, { "職", ["ək"]}, { "蒸", ["əŋ"]},
            { "支", ["ɛ"]}, { "錫", ["ɛk"]}, { "耕", ["ɛŋ"]},
            { "侯", ["ɔ"]}, { "屋", ["ɔk"]}, { "東", ["ɔŋ"]},
            { "幽", ["o"]}, { "覺", ["ok"]}, { "冬", ["oŋ"]},
            { "宵", ["ø"]}, { "藥", ["øk"]}, 
            { "歌", ["al", "ɛl"]},*
            { "月", ["at", ["ɛt"]]}, { "質", ["et"]}, { "元", ["an", "ɔn", "ɛn"]}, */{ "物", ["ət", "ot"]}
            /*  { "微", ["əl"]},
                { "脂", ["el"]},

                */
        };

        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Workbook wk = new Workbook("D:/shang4gu3li3in1NoOt.xlsx");
            Worksheet ws = wk.Worksheets[0];
            //CheckDen(ws);
            int length = CheckDoubleMapping(ws);

            Workbook wbForSave = new Workbook();
            int sheetNr = 0;
            foreach (var k in vin14bu4.Keys)
            {
                string vin11 = k;
                string vin12 = k;
                System.Threading.Thread.Sleep(5000);
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + vin11 + "&y=" + vin12 + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, wbForSave, sheetNr, length, vin11, vin12);
                wbForSave.Worksheets.Add();
                sheetNr ++;
            }
            string fn = vin14bu4.Keys.Count>3 ? "shang4gu3vin4jo5" : string.Concat(vin14bu4.Keys.AsEnumerable());
            wbForSave.Save(@"D:\" + fn + DateTime.Now.ToString("yyMMddHHmm") + ".xlsx");
        }

        private static void ProcessTable(string theText, Worksheet ws, Workbook wbForSave, int sheetNr, int length, string vin11, string vin12)
        {
            try { 
            Worksheet ws2 = wbForSave.Worksheets[sheetNr];
            var dt2 = ws2.Cells.ExportDataTable(0, 0, 1600, 9);
            string[] lines = theText.Split( new string[] { Environment.NewLine },    StringSplitOptions.None);
            string table = lines.Where(x=>x.StartsWith("<table><tr><th")).FirstOrDefault();
            if (table != null) { 
                lines = table.Split(new string[] { "<tr><td>" }, StringSplitOptions.None);
                int hang2 = 0;
                int miou4su4 = 0;
                int chu5vin4zy4su4 = 0, cy3bu4zy4su4 = 0;
                string chu5vin4zy4 = "", vin4jo5zy4 = "";
                foreach (string line in lines)
                {
                    var rythms = line.Split(new string[] { "<b style=\"" }, StringSplitOptions.None);
                    List<string> vals = new List<string>();
                    if (rythms.Length > 1)
                    {
                        for (int i = 0; i < rythms.Length - 1; i++)
                        {
                            string zy = rythms[i].Substring(rythms[i].Length - 1);
                            if (zy == "A" || zy == "B")
                                zy = rythms[i].Substring(rythms[i].Length - 2, 1);
                            if (zy.Contains("\ude62") || zy.Contains("\udfae"))
                            {
                                zy = rythms[i].Substring(rythms[i].Length - 2, 2);
                            }
                            if (zy.Contains("}"))
                            {
                                zy = rythms[i].Substring(rythms[i].Length - 3, 3);
                            }
                            if (zy.Contains("\udf1a"))
                            {
                                zy = "盧"; //白公父簠金文是盧
                            }
                            if (zy.Contains("𫭠"))
                            {
                                zy = "筐"; //史免簠金文是⿷匚𫭠
                            }
                            else if (zy.Contains("宮九"))
                            {
                                zy = "九"; //叔卣金文是⿰宮九
                            }
                            else if (zy == "卝")
                            {
                                 zy = "丱";
                            }
                            //if (zy == "丱")
                            //{
                            //    int x = 0;
                            //}
                            Console.Write(zy);
                            CalcTotalHanzy(zy, ref vin4jo5zy4, ref cy3bu4zy4su4);
                            vals.Add(zy);
                            Dictionary<string, string> do1in1 = new Dictionary<string, string>();
                            //List<string> zhong4gu3do1in1 = new List<string>();
                            for (int j = 1; j < length; j++)              
                            {
                                if (ws.Cells["O" + j.ToString()].Value == null || ws.Cells["G" + j.ToString()].Value == null)//"O"列是同聲旁同音字"G"列是上古音
                                        continue;
                                else if (ws.Cells["O" + j.ToString()].Value.ToString().Contains(zy) && 
                                    ws.Cells["O" + j.ToString()].GetStyle().Font.Color != System.Drawing.ColorTranslator.FromHtml("#ffffcc00"))//"O"列是同聲旁同音字
                                    {                                
                                        var shang4gu3du5in1 = ws.Cells["G" + j.ToString()].Value.ToString(); //"G"列是上古音
                                        var zhong1gu3du5in1 = ws.Cells["N" + j.ToString()].Value.ToString(); //"N"列是中古音
                                        if (!do1in1.Keys.Contains(shang4gu3du5in1)) 
                                            do1in1.Add(shang4gu3du5in1, zhong1gu3du5in1);
                                        //zhong4gu3do1in1.Add(zhong4gu3du5in1);
                                        if (vin14bu4[vin11].All(d => !shang4gu3du5in1.Contains(d))
                                         || (vin11 == "之" && shang4gu3du5in1.Contains("əl"))
                                         || (vin11 == "幽" && (shang4gu3du5in1.EndsWith("l") || shang4gu3du5in1.EndsWith("lh") || shang4gu3du5in1.EndsWith("lɣ"))))
                                        {
                                            shang4gu3du5in1 += "謬";
                                            miou4su4++;
                                        }
                                        Console.Write(shang4gu3du5in1 + "/");
                                        vals.Add(shang4gu3du5in1);
                                    }
                            }
                            if (do1in1.All(x => vin14bu4[vin11].All(d => !x.Key.Contains(d))))
                            {//多音字所有音都不合韻部，先嘗試人工智能修正，正不了再確定出韻
                                var do1in1Copy = do1in1;
                                bool i3siou1zhen4 = false;
                                foreach (var gu3in1 in do1in1Copy.ToList())
                                {
                                    string qi2ta1shang4gu3in1 = FindRightOldPronunciation(ws, length, vin14bu4[vin11], gu3in1.Value, gu3in1.Key);
                                    if (qi2ta1shang4gu3in1 != gu3in1.Key)
                                    {
                                        do1in1Copy.Add(qi2ta1shang4gu3in1, gu3in1.Value);
                                        do1in1Copy.Remove(gu3in1.Key);
                                        i3siou1zhen4 = true;
                                    }                                    
                                }
                                if (!i3siou1zhen4)
                                {
                                    CalcTotalHanzy(zy, ref chu5vin4zy4, ref chu5vin4zy4su4);
                                    vals[vals.IndexOf(zy)] += "謬";
                                }
                            }
                        }
                        Console.WriteLine();
                        int lie5 = 0;
                        foreach (var v in vals)
                        {                        
                            if (v.EndsWith("謬"))
                            {
                                Style style = new Style();
                                style.Font.Color = Color.Red;
                                ws2.Cells[hang2, lie5].SetStyle(style);
                            }
                            ws2.Cells[hang2, lie5].Value = v.Replace("謬", "");
                            lie5++;
                        }
                        hang2++;
                    }
                }
                string chu5vin4zy4tong3ji4 = vin12 + "部紅色出韻字" + chu5vin4zy4su4.ToString() + "個：" + chu5vin4zy4;
                ws2.Cells[hang2, 0].Value = vin12 + "部紅色出韻音" + miou4su4.ToString() + "個";
                ws2.Cells[hang2, 8].Value = chu5vin4zy4tong3ji4;
                ws2.Cells[hang2 +1, 0].Value = vin12 + "部韻腳字" + cy3bu4zy4su4.ToString() + "個: " + vin4jo5zy4;
                Console.Write(chu5vin4zy4tong3ji4);
            }
            else
            {
                int sas = 0;
            }
            }
            catch (Exception e)
            {
                int sas = 0;
            }
        }

        private static void CalcTotalHanzy(string zy, ref string so3iou3zy4, ref int su4liang4)
        {
            if (!so3iou3zy4.Contains(zy))
            {
                su4liang4++;
                so3iou3zy4 += zy;
            }
        }
        private static string FindRightOldPronunciation(Worksheet ws, int length, string[] vin4bu4, string zhong1gu3in1, string shang4gu3in1)
        {
            string res = shang4gu3in1;
            try { 
                for (int j = 1; j < length; j++)
                {
                    if (ws.Cells["N" + j.ToString()].Value != null && //"N"列是中古音
                        ws.Cells["G" + j.ToString()].Value != null && //"G"列是上古音
                        ws.Cells["N" + j.ToString()].Value.Equals(zhong1gu3in1) &&
                        vin4bu4.Any(x => ws.Cells["G" + j.ToString()].Value.ToString().Contains(x))
                        )
                    {
                        res = ws.Cells["G" + j.ToString()].Value.ToString();
                        return res;
                    }
                }
            }
            catch (Exception e)
            { 
                return res; 
            }
            return res;
        }

        private static int CheckDoubleMapping(Worksheet ws)
        {
            Dictionary<string, List<string>> Mapping = new Dictionary<string, List<string>>();
            int res = 3;
            int nullcount = 0;
            while (ws.Cells["G" + res.ToString()].Value == null || 
                !String.IsNullOrWhiteSpace(ws.Cells["G" + res.ToString()].Value.ToString())) //"G"列是上古音
            {
                if (ws.Cells["G" + res.ToString()].Value != null)//"G"列是上古音
                {
                    string k = ws.Cells["G" + res.ToString()].Value.ToString();//"G"列是上古音
                    if (ws.Cells["N" + res.ToString()].Value != null)//"N"列是中古音
                    { 
                        string v = ws.Cells["N" + res.ToString()].Value.ToString()+ ws.Cells["K" + res.ToString()].Value.ToString();//"K"列是中古韻
                        if (!Mapping.ContainsKey(k))
                        {
                            Mapping.Add(k, new List<string>());
                        }
                        if (!Mapping[k].Contains(v))
                        {
                            Mapping[k].Add(v);
                        }
                        nullcount = 0;
                    }
                }
                else
                {
                    nullcount++;
                }
                if (nullcount > 4)
                    break;
                res++;
            }
            foreach (var kv in Mapping)
            {
                if (kv.Value.Count>1)
                {
                    Console.Write(kv.Key + " ");
                    foreach (var v in kv.Value)
                        Console.Write(v + " ");
                    Console.WriteLine();
                }
            }
            return res - nullcount;

        }

        private static void CheckDen(Worksheet ws)
        {
            Dictionary<string, List<string>> Siao1di5den3 = new Dictionary<string, List<string>>() {
                { "一", new List<string>() },{ "二", new List<string>() },{ "三", new List<string>() },{ "四", new List<string>() },
            };
            int res = 3;
            while (ws.Cells["G" + res.ToString()].Value == null ||
                !String.IsNullOrWhiteSpace(ws.Cells["G" + res.ToString()].Value.ToString())) //"G"列是上古音
            {
                if (ws.Cells["G" + res.ToString()].Value != null)//"G"列是上古音
                {
                    string oe = ws.Cells["G" + res.ToString()].Value.ToString();//"G"列是上古音
                    if (oe.Contains("ø") && !oe.Contains("øk"))
                    {
                        string k = ws.Cells["I" + res.ToString()].Value.ToString();//"I"列是中古等
                        string v = ws.Cells["K" + res.ToString()].Value.ToString();//"I"列是中古韻
                        if (!Siao1di5den3[k].Contains(v))
                        {
                            Siao1di5den3[k].Add(v);
                        }
                    }
                }
                res++;
            }
            foreach (var kv in Siao1di5den3)
            {
                if (kv.Value.Count > 1)
                {
                    Console.Write(kv.Key + " ");
                    foreach (var v in kv.Value)
                        Console.Write(v + " ");
                    Console.WriteLine();
                }
            }
        }

        
    }    
}