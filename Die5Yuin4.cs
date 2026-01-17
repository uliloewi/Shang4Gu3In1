using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace Shang4Gu3In1
{
    /// <summary>
    ///    檢查連綿詞擬音阿合理
    /// </summary>
    internal class Die5Yuin4
    {
        public static void Main(string[] args)
        {
            const string uen2jän4ja5 = @"C:\Users\xggg\Downloads\SieShu\";
            Console.OutputEncoding = Encoding.UTF8;
            Workbook wk = new Workbook(uen2jän4ja5 + "a.xlsx");
            Worksheet ws = wk.Worksheets[0];
            foreach (string line in File.ReadLines(uen2jän4ja5 + "連綿詞.txt"))
            {
                //var a = GetWord(line,  "：", " ");
                //Console.WriteLine(a);
                //var a = GetWord(line, "（", "）");
                //var wd = line.Replace("（" + a + "）", "");
                //Console.WriteLine(wd);
                var x = line.Split("/");
                foreach (string s in x)
                {
                    if (!Regex.IsMatch(s, "[a-zA-Z]"))
                    {
                        string a = "";
                        string b = "";
                        switch (s.Length)
                        {
                            case 2:
                                a = s[0].ToString();
                                b = s[1].ToString();
                                NeedNotice(a, b, ws);
                                break;
                            case 3:
                                a = s[0].ToString();
                                b = s.Substring(1);

                                NeedNotice(a, b, ws);

                                a = s.Substring(0, 2);
                                b = s.Substring(2);
                                NeedNotice(a, b, ws);
                                break;
                            case 4:
                                a = s.Substring(0, 2);
                                b = s.Substring(2, 2);
                                NeedNotice(a, b, ws);
                                break;
                            default:
                                break;
                        }
                    }

                }
            }
        }

        static void NeedNotice(string a, string b, Worksheet ws)
        {
            if (a == b) return;
            bool res = false;
            Dictionary<string, List<string>> map = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };
            Dictionary<string, List<string>> sheng1 = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };
            Dictionary<string, List<string>> in1zie5 = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };

            for (int row = 3; row <= ws.Cells.MaxDataRow; row++)
            {//處理了一字所有古音
                if (ws.Cells["P" + row.ToString()].Value == null || ws.Cells["G" + row.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                    continue;
                if (ws.Cells["P" + row.ToString()].Value.ToString().Contains(a))//"P"列是同聲旁同音字
                {
                    map[a].Add(ws.Cells["E" + row.ToString()].Value?.ToString());
                    sheng1[a].Add(ws.Cells["H" + row.ToString()].Value?.ToString());
                    in1zie5[a].Add(ws.Cells["G" + row.ToString()].Value?.ToString());
                }
                if (ws.Cells["P" + row.ToString()].Value.ToString().Contains(b))//"P"列是同聲旁同音字
                {
                    map[b].Add(ws.Cells["E" + row.ToString()].Value?.ToString());
                    sheng1[b].Add(ws.Cells["H" + row.ToString()].Value?.ToString());
                    in1zie5[b].Add(ws.Cells["G" + row.ToString()].Value?.ToString());
                }
            }

            Console.Write(a);
            Console.Write(string.Join(";", in1zie5[a]));
            Console.Write(b);
            Console.WriteLine(string.Join(";", in1zie5[b]));
        }


        static bool NeedNotice2(string a, string b, Worksheet ws)
        {
            if (a == b) return false;
            bool res = false;
            Dictionary<string, List<string>> map = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };
            Dictionary<string, List<string>> sheng1 = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };
            Dictionary<string, List<string>> in1zie5 = new Dictionary<string, List<string>>() { { a, new List<string>() }, { b, new List<string>() } };

            for (int row = 3; row <= ws.Cells.MaxDataRow; row++)
            {//處理了一字所有古音
                if (ws.Cells["P" + row.ToString()].Value == null || ws.Cells["G" + row.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                    continue;
                if (ws.Cells["P" + row.ToString()].Value.ToString().Contains(a))//"P"列是同聲旁同音字
                {
                    map[a].Add(ws.Cells["E" + row.ToString()].Value?.ToString());
                    sheng1[a].Add(ws.Cells["H" + row.ToString()].Value?.ToString());
                    in1zie5[a].Add(ws.Cells["G" + row.ToString()].Value?.ToString());
                }
                if (ws.Cells["P" + row.ToString()].Value.ToString().Contains(b))//"P"列是同聲旁同音字
                {
                    map[b].Add(ws.Cells["E" + row.ToString()].Value?.ToString());
                    sheng1[b].Add(ws.Cells["H" + row.ToString()].Value?.ToString());
                    in1zie5[b].Add(ws.Cells["G" + row.ToString()].Value?.ToString());
                }
            }
            //res = map[a].Count==0 || map[b].Count==0 || map[a].All(x=>!map[b].Contains(x));
            if (sheng1[a].Count > 0 && sheng1[b].Count > 0 && sheng1[a].Any(x => sheng1[b].Contains(x))//雙聲
                && in1zie5[a].All(x => !x.StartsWith("s")) && in1zie5[b].All(x => !x.StartsWith("s")))//不是sK,sT
                return false;
            res = map[a].Count == 0 || map[b].Count == 0 || map[a].All(x => !map[b].Contains(x));

            if (!res && in1zie5[a].Count > 0 && in1zie5[b].Count > 0)
            {
                Console.WriteLine(a);
                foreach (var x in in1zie5[a])
                    Console.WriteLine(x);
                Console.WriteLine(b);
                foreach (var x in in1zie5[b])
                    Console.WriteLine(x);
            }

            return res;
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
    }
}
