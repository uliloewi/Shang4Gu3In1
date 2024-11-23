using Shang4Gu3In1;
using Aspose.Cells;
using System.Drawing;

namespace zhongguliin
{
    class Program
    {
        private static Dictionary<string, string> uin4bu4 = new Dictionary<string, string>() {
                { "魚", "a" },
                { "鐸", "ak" },
                { "蒸", "əŋ" },
                { "脂", "el" },
            };

        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            string uin1 = "蒸";
            string uin2 = "蒸";
            var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + uin1 + "&y=" + uin2 + "&mode=yunbu");
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            Workbook wk = new Workbook("./shang4gu3li3in1.xlsx");
            Worksheet ws = wk.Worksheets[0];
            ProcessTable(content, ws, uin1, uin2);
        }

        private static void ProcessTable(string theText, Worksheet ws, string uin1, string uin2)
        {
            Workbook wb2 = new Workbook();
            Worksheet ws2 = wb2.Worksheets[0];
            var dt2 = ws2.Cells.ExportDataTable(0, 0, 1600, 9);
            string[] lines = theText.Split(    new string[] { Environment.NewLine },    StringSplitOptions.None);
            string table = lines.Where(x=>x.StartsWith("<table><tr><th")).FirstOrDefault();
            lines = table.Split(new string[] { "<tr><td>" }, StringSplitOptions.None);
            int hang2 = 0;
            foreach (string line in lines)
            {
                var rythms = line.Split(new string[] { "<b style=\"" }, StringSplitOptions.None);
                List<string> vals = new List<string>();
                if (rythms.Length > 1)
                {
                    for (int i = 0; i < rythms.Length - 1; i++)
                    {
                        string zy = rythms[i].Substring(rythms[i].Length - 1);
                        Console.Write(zy);
                        vals.Add(zy);
                        for (int j = 1; j < 9915; j++)              
                        {
                            if (ws.Cells["L" + j.ToString()].Value.ToString().Contains(zy) && ws.Cells["L" + j.ToString()].GetStyle().Font.Color != System.Drawing.ColorTranslator.FromHtml("#ffffcc00"))
                            {                                
                                var du5in1 = ws.Cells["D" + j.ToString()].Value.ToString();
                                if (!du5in1.Contains(uin4bu4[uin1]))
                                {
                                    du5in1 += "謬";
                                }
                                Console.Write(du5in1 + "/");
                                vals.Add(du5in1);
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
            wb2.Save(@"D:\"+ uin1 + uin2 + ".xlsx");
        }
    }    
}