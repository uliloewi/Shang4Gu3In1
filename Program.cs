using Shang4Gu3In1;
using Aspose.Cells;
using System.Text;

namespace zhongguliin
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=%E6%A5%9A%E8%BE%AD%E9%9F%BB&x=%E9%AD%9A&y=%E9%AD%9A&mode=yunbu");
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            Workbook wk = new Workbook("./shang4gu3li3in1.xlsx");
            Worksheet ws = wk.Worksheets[0];
            //var dt = ws.Cells.ExportDataTable(0, 0, 10000, 15);
            //var shangguin = ws.Cells["D1"].Value.ToString();
            ProcessTable(content, ws);
        }

        private static void ProcessTable(string theText, Worksheet ws)
        {
            string[] lines = theText.Split(    new string[] { Environment.NewLine },    StringSplitOptions.None);
            string table = lines.Where(x=>x.StartsWith("<table><tr><th")).FirstOrDefault();
            lines = table.Split(new string[] { "<tr><td>" }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                var rythms = line.Split(new string[] { "<b style=\"" }, StringSplitOptions.None);
                if (rythms.Length > 1)
                {

                    for (int i = 0; i < rythms.Length-1; i++)
                    {
                        string zy = rythms[i].Substring(rythms[i].Length - 1);
                        Console.Write(zy);
                        for (int j = 1; j < 9915; j++)
                        {
                            if (ws.Cells["L" + j.ToString()].Value.ToString().Contains(zy))
                                Console.Write(ws.Cells["D" + j.ToString()].Value.ToString() + "/");
                        }
                    }
                    Console.WriteLine();
                }
            } 
        }
    }    
}