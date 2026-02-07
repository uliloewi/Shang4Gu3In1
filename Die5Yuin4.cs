using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
namespace Shang4Gu3In1
{
    /// <summary>
    ///    檢查連綿詞擬音阿合理
    /// </summary>
    public class Die5Yuin4
    {
        public static void Main1(string[] args, string uen2jän4ja5)
        {

            //const string uen2jän4ja5 = @"C:\Users\xggg\Downloads\SieShu\";
            Console.OutputEncoding = Encoding.UTF8;
            Workbook wk = new Workbook(uen2jän4ja5 + "a.xlsx");
            Worksheet ws = wk.Worksheets[0];
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
            }
            /*
            foreach (string line in File.ReadLines(uen2jän4ja5 + "連綿詞.txt"))
            {
                var x = line.Split("/");
                foreach (string s in x)
                {
                    if (!Regex.IsMatch(s, "[a-zA-Z]"))
                    {
                        NeedNotice(s, ws);
                    }
                }
            }
            */
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

        static void NeedNotice(string zy, Worksheet ws)
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
            if (!AllListsShareCommonElement(in1zie5, true))
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

            IEnumerable<string> intersection = dic.First().Value;

            // Wir vergleichen die Basis nacheinander mit jeder weiteren Liste
            for (int i = 1; i < dic.Count; i++)
            {
                intersection = intersection.Intersect(dic.ElementAt(i).Value);
            }

            // Wenn am Ende noch Elemente übrig sind, gibt es ein gemeinsames Element in ALLEN Listen
            return intersection.Any();
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