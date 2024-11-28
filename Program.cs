using Shang4Gu3In1;
using Aspose.Cells;
using System.Drawing;

namespace zhongguliin
{
    class Program
    {
        private static Dictionary<string, string> uin4bu4 = new Dictionary<string, string>() {
                /*{ "魚", "a" }, { "鐸", "ak" }, { "陽", "aŋ" },
                { "之", "ə" }, { "職", "ək" }, { "蒸", "əŋ" },
                { "支", "ɛ" },  { "錫", "ɛk" }, { "耕", "ɛŋ" },
                { "侯", "ɔ" }, { "屋", "ɔk" }, { "東", "ɔŋ" },
                { "幽", "o" }, { "覺", "ok" }, { "冬", "oŋ" },*/
                { "宵", "ɔl" }, { "藥", "ɔlk" },
             /*   { "宵", "ø" }, { "藥", "øk" }, 
            /*  { "微", "əl" },
                { "脂", "el" },
                
                */
            };

        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Workbook wk = new Workbook("D:/shang4gu3li3in1.xlsx");
            Worksheet ws = wk.Worksheets[0];
            //CheckDen(ws);
            int length = CheckDoubleMapping(ws);
            foreach (var k in uin4bu4.Keys)
            {
                string uin1 = k;
                string uin2 = k;
                System.Threading.Thread.Sleep(5000);
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + uin1 + "&y=" + uin2 + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, length, uin1, uin2);
            }
        }

        private static void ProcessTable(string theText, Worksheet ws, int length, string uin1, string uin2)
        {
            string siao1io5 = "宵藥";
            Workbook wb2 = new Workbook();
            Worksheet ws2 = wb2.Worksheets[0];
            var dt2 = ws2.Cells.ExportDataTable(0, 0, 1600, 9);
            string[] lines = theText.Split(    new string[] { Environment.NewLine },    StringSplitOptions.None);
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
                            if (zy.Contains("\ude62"))
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
                            //if (zy == "沃")
                            //{
                            //    int x = 0;
                            //}
                            Console.Write(zy);
                            StaticNumer(zy, ref vin4jo5zy4, ref cy3bu4zy4su4);
                            vals.Add(zy);
                            List<string> do1in1 = new List<string>();
                            for (int j = 1; j < length; j++)              
                            {
                                if (ws.Cells["L" + j.ToString()].Value.ToString().Contains(zy) && ws.Cells["L" + j.ToString()].GetStyle().Font.Color != System.Drawing.ColorTranslator.FromHtml("#ffffcc00"))
                                {                                
                                    var du5in1 = ws.Cells["D" + j.ToString()].Value.ToString();
                                    do1in1.Add(du5in1);
                                  /*  if (!du5in1.Contains(uin4bu4[uin1])
                                     || (uin1 == "之" && du5in1.Contains("əl"))
                                     || (uin1 == "幽" && (du5in1.EndsWith("l") || du5in1.EndsWith("lh") || du5in1.EndsWith("lɣ"))))*/
                                   if ((uin1 == "藥" && !du5in1.Contains("olk") && !du5in1.Contains("ɔlk"))
                                        || (uin1 == "宵" && !du5in1.Contains("ol") && !du5in1.Contains("ɔl"))
                                        || (!"宵藥".Contains(uin1) && !du5in1.Contains(uin4bu4[uin1]))
                                        || (uin1 == "之" && du5in1.Contains("əl"))
                                        || (uin1 == "幽" && (du5in1.EndsWith("l") || du5in1.EndsWith("lh") || du5in1.EndsWith("lɣ"))))
                                    {
                                        du5in1 += "謬";
                                        miou4su4++;
                                    }
                                    Console.Write(du5in1 + "/");
                                    vals.Add(du5in1);
                                }
                            }
                           // if (do1in1.All(x => !x.Contains(uin4bu4[uin1])))
                            if ((uin1 == "藥" & do1in1.All(x => !x.Contains("olk") && !x.Contains("ɔlk"))) ||
                                (uin1 == "宵" && do1in1.All(x => !x.Contains("ol") && !x.Contains("ɔl"))) ||
                                (!"宵藥".Contains(uin1) && do1in1.All(x => !x.Contains(uin4bu4[uin1]))))
                            {
                                StaticNumer(zy, ref chu5vin4zy4, ref chu5vin4zy4su4);
                                vals[vals.IndexOf(zy)] += "謬";
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
                string chu5vin4zy4tong3ji4 = "紅色出韻字" + chu5vin4zy4su4.ToString() + "個：" + chu5vin4zy4;
                ws2.Cells[hang2, 0].Value = "紅色出韻音" + miou4su4.ToString() + "個";
                ws2.Cells[hang2, 8].Value = chu5vin4zy4tong3ji4;
                ws2.Cells[hang2 +1, 0].Value = "韻腳字" + cy3bu4zy4su4.ToString() + "個: " + vin4jo5zy4;
                Console.Write(chu5vin4zy4tong3ji4);
                wb2.Save(@"D:\"+ uin1 + uin2 + ".xlsx");
            }
            else
            {
                int sas = 0;
            }
        }

        private static void StaticNumer(string zy, ref string so3iou3zy4, ref int su4liang4)
        {
            if (!so3iou3zy4.Contains(zy))
            {
                su4liang4++;
                so3iou3zy4 += zy;
            }
        }

        private static int CheckDoubleMapping(Worksheet ws)
        {
            Dictionary<string, List<string>> Mapping = new Dictionary<string, List<string>>();
            int res = 3;
            int nullcount = 0;
            while (ws.Cells["D" + res.ToString()].Value == null || 
                !String.IsNullOrWhiteSpace(ws.Cells["D" + res.ToString()].Value.ToString()))
            {
                if (ws.Cells["D" + res.ToString()].Value != null)
                {
                    string k = ws.Cells["D" + res.ToString()].Value.ToString();
                    if (ws.Cells["K" + res.ToString()].Value != null)
                    { 
                        string v = ws.Cells["K" + res.ToString()].Value.ToString()+ ws.Cells["H" + res.ToString()].Value.ToString();
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
            return res - 1;

        }

        private static void CheckDen(Worksheet ws)
        {
            Dictionary<string, List<string>> Siao1di5den3 = new Dictionary<string, List<string>>() {
                { "一", new List<string>() },{ "二", new List<string>() },{ "三", new List<string>() },{ "四", new List<string>() },
            };
            int res = 3;
            while (ws.Cells["D" + res.ToString()].Value == null ||
                !String.IsNullOrWhiteSpace(ws.Cells["D" + res.ToString()].Value.ToString()))
            {
                if (ws.Cells["D" + res.ToString()].Value != null)
                {
                    string oe = ws.Cells["D" + res.ToString()].Value.ToString();
                    if (oe.Contains("ø") && !oe.Contains("øk"))
                    {
                        string k = ws.Cells["F" + res.ToString()].Value.ToString();
                        string v = ws.Cells["H" + res.ToString()].Value.ToString();
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