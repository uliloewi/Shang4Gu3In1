using Shang4Gu3In1;
using Aspose.Cells;
using System.Text;

namespace zhongguliin
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=%E6%A5%9A%E8%BE%AD%E9%9F%BB&x=%E9%AD%9A&y=%E9%AD%9A&mode=yunbu");
            var content = await httpResponseMessage.Content.ReadAsStringAsync();
            ProcessTable(content);
        }

        private static void ProcessTable(string theText)
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
                    }
                    Console.WriteLine();
                }
            } 
        }
    }    
}