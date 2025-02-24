using Shang4Gu3In1;
using Aspose.Cells;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace Shang4Gu3In1
{
    class Program
    {


        const string lu5vwn2in1 = "aɛɔəøo";//六元音
        /*
              { "鐸", "ak"}, { "錫", "ɛk"}, { "屋", "ɔk"}, { "職", "ək"}, { "藥", "øk"}, { "覺", "ok"},
              { "陽", "aŋ"}, { "耕", "ɛŋ"}, { "東", "ɔŋ"}, { "蒸", "əŋ"}, { "冬", "oŋ"},
              { "歌", "al", "ɛl"},  { "微", "əl"}, { "脂", "øl"},
              { "月", "at", （拜）"ɛt"}, { "物", "ət"}, { "質", "øt"},
              { "元", （山）"an", （見）"ɛn", "ɔn"}, { "文", "ən"},
              { "真", "øn", （印）"øŋ"},
              { "葉", "ap", （怗）"øp"}, { "緝", "əp"},
              { "談", "am"}, { "侵", "əm", （風）"om", （添）"øm"},
              { "魚", "a"}, { "之", "ə"}, { "支", "ɛ"},
              { "侯", "ɔ"}, { "幽", "o"}, { "宵", "ø"},
         */

        private static Dictionary<string, string[]> shang4gu3vin4bu4 = new Dictionary<string, string[]>() {//三十一韻部
             { "鐸", [lu5vwn2in1[0]+"k"]}, { "錫", [lu5vwn2in1[1]+"k"]}, { "屋", [lu5vwn2in1[2]+"k"]}, { "職", [lu5vwn2in1[3]+"k"]}, { "藥", [lu5vwn2in1[4]+"k"]}, { "覺", [lu5vwn2in1[5]+"k"]},
             { "陽", [lu5vwn2in1[0]+"ŋ"]}, { "耕", [lu5vwn2in1[1]+"ŋ"]}, { "東", [lu5vwn2in1[2]+"ŋ"]}, { "蒸", [lu5vwn2in1[3]+"ŋ"]}, { "冬", [lu5vwn2in1[5]+"ŋ"]},
             { "歌", [lu5vwn2in1[0]+"l", lu5vwn2in1[1]+"l"]},  { "微", [lu5vwn2in1[3]+"l"]}, { "脂", [lu5vwn2in1[4]+"l"]},
             { "月", [lu5vwn2in1[0]+"t", lu5vwn2in1[1]+"t"]}, { "物", [lu5vwn2in1[3]+"t"]}, { "質", [lu5vwn2in1[4]+"t"]},
             { "元", [lu5vwn2in1[0]+"n", lu5vwn2in1[1]+"n", lu5vwn2in1[2]+"n"]}, { "文", [lu5vwn2in1[3]+"n"]},
             { "真", [lu5vwn2in1[4]+"n", lu5vwn2in1[4]+"ŋ"]},
             { "葉", [lu5vwn2in1[0]+"p", lu5vwn2in1[4]+"p"]}, { "緝", [lu5vwn2in1[3]+"p"]},
             { "談", [lu5vwn2in1[0]+"m"]}, { "侵", [lu5vwn2in1[3]+"m", lu5vwn2in1[4]+"m", lu5vwn2in1[5]+"m"]},
             { "魚", [lu5vwn2in1[0].ToString()]}, { "之", [lu5vwn2in1[3].ToString()]}, { "支", [lu5vwn2in1[1].ToString()]},
             { "侯", [lu5vwn2in1[2].ToString()]}, { "幽", [lu5vwn2in1[5].ToString()]}, { "宵", [lu5vwn2in1[4].ToString()]},
        };

        private static Dictionary<string, int> vin4bu4hao4 = new Dictionary<string, int>() {//三十一韻部
             {lu5vwn2in1[0]+"k", 1}, { lu5vwn2in1[1]+"k", 2}, { lu5vwn2in1[2]+"k",3}, { lu5vwn2in1[3]+"k", 4}, { lu5vwn2in1[4]+"k",5}, { lu5vwn2in1[5]+"k",6},
             { lu5vwn2in1[0]+"ŋ", 7}, { lu5vwn2in1[1]+"ŋ", 8}, { lu5vwn2in1[2]+"ŋ", 9}, { lu5vwn2in1[3]+"ŋ", 10}, { lu5vwn2in1[5]+"ŋ", 11},
             { lu5vwn2in1[0]+"l",12 }, {lu5vwn2in1[1]+"l",13},  { lu5vwn2in1[3]+"l",14}, { lu5vwn2in1[4]+"l",15},
             { lu5vwn2in1[0]+"t",16 },{ lu5vwn2in1[1]+"t",17}, { lu5vwn2in1[3]+"t",18}, { lu5vwn2in1[4]+"t", 19},
             { lu5vwn2in1[0]+"n", 20 },{ lu5vwn2in1[1]+"n", 21 }, {lu5vwn2in1[2]+"n",22}, { lu5vwn2in1[3]+"n",23},
             { lu5vwn2in1[4]+"n", 24 },{ lu5vwn2in1[4]+"ŋ", 25},
             { lu5vwn2in1[0]+"p", 25 },{lu5vwn2in1[4]+"p",27}, { lu5vwn2in1[3]+"p",28},
             { lu5vwn2in1[0]+"m", 29}, { lu5vwn2in1[3]+"m", 30 },{lu5vwn2in1[4]+"m", 31 },{lu5vwn2in1[5]+"m", 32},
             { lu5vwn2in1[0].ToString(), 33}, { lu5vwn2in1[3].ToString(), 34}, { lu5vwn2in1[1].ToString(), 35},
             { lu5vwn2in1[2].ToString(), 36}, { lu5vwn2in1[5].ToString(), 37}, { lu5vwn2in1[4].ToString(), 38},
        };

        private static Dictionary<string, string[]> shen1mu3duei4in4 = new Dictionary<string, string[]>() {//上古中古聲母對映
            { "sKr莊組A", ["skr", "skʰr", "sgr", "sxr", "sɣr"]},            
            { "sK精組B", ["skʰ", "sk", "sg", "sx", "sɣ"]},  
            { "sC心C", ["sn", "sŋ", "sm"]}, 
            { "sl心母D", ["sl"]}, 
            { "st心書母E", ["st"]}, 
            { "sd從邪母F", ["sd"]},
            { "rT端組日母G", ["rtʰ", "rt", "rd", "rn"]}, 
            { "xm曉H", ["xm"]},
            { "sr生母I", ["sr"]},
            { "Kl章端組以母J", ["kl", "kʰl", "gl", "ŋl", "xl", "xn", "ɣl", "l"]},
            { "Tr知組K", ["tr", "tʰr", "dr", "nr"]}, 
            { "rK知組L",["rkʰ", "rk", "rg", "rx", "rŋ"]},
            { "r來母M", ["rɣ", "r"]}, 
            { "N明泥日娘組N", ["n"]}, 
            { "T章端組O", [ "tʰ", "t", "d"]}, 
            { "M明母P", ["m"]}, 
            { "K見溪群疑Q", ["kʰ", "k",  "g", "ŋ"]}, 
            { "X曉匣R", ["x", "ɣ"]}, // { "rɣl知組",["rɣl"]},
            { "P幫組S", ["pʰ", "p", "b"]},
            { "h影母T", ["h"]},
        };

        static List<string> tong1ia5 = new List<string>() { "魚鐸", "魚陽", "魚之", "魚支", "魚侯", "魚屋", "魚東", "魚幽", "魚宵", "魚歌", "魚元", "魚微",
            "鐸之", "鐸職", "鐸錫", "鐸侯", "鐸屋", "鐸幽", "鐸藥", "鐸歌", "鐸質", "鐸葉", "陽蒸", "陽錫", "陽耕", "陽東", "陽冬", 
            "之職", "之蒸", "之侯", "之幽", "之覺", "之宵", "之元", "之緝", "職蒸", "職侯", "職屋", "職幽", "職覺", "職藥", "職葉", "職緝", 
            "蒸侯", "蒸東", "蒸冬", "蒸文", "蒸侵", "支錫", "支歌", "支月", "支元", "支微", "支物", "支文", "支脂", "支質",
            "錫屋", "錫歌", "錫月", "錫物", "錫質", "耕東", "耕元", "耕文", "耕真", "侯屋", "侯東", "侯幽", "侯冬", "侯宵", 
            "屋覺", "屋宵", "屋藥", "東幽", "東冬", "東侵", "幽覺", "幽宵", "覺宵", "覺緝", "冬真", "冬侵", "宵元", "歌元", "歌微", "歌物", "歌脂", 
            "月元", "月物", "月脂", "月質", "月葉", "月緝", "元微", "元物", "元文", "元脂", "元質", "元真", "元緝", "微物", "微文", "微脂", "微質", 
            "物脂", "物質", "文質", "文真", "文緝", "脂質", "質真", "葉談", "葉緝", "談侵"};

        private static string zhong1gu3vwn2in1 = "aeiouvwryäüöëï";//廣通中古拼音的元音

        private static Dictionary<string, string> shen1pang2vin4luei4 = new Dictionary<string, string>();

        static async Task Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Workbook wk = new Workbook("D:/《廣韻》形聲考李.xlsx");
            Worksheet ws = wk.Worksheets[0];
            //CheckDen(ws);
            int length = CheckDoubleMapping(ws);
            Dictionary<string, int> d = OnsetsOC(ws, length);
            foreach (var s in d.OrderBy(x=>x.Value))
            {
                Console.WriteLine(s);
            }
           /* foreach (var s in GetPhoneticComponent(ws,length))
            {
                Console.WriteLine(s);
            }*/            
            foreach (var s in shen1pang2vin4luei4)
            {
                Console.WriteLine(s);
            }
            //Console.WriteLine(lu5vwn2in1[1]);            
            Workbook wbForSave = new Workbook();
            shang4gu3duei4zhong1gu3(ws, length);
            int sheetNr = 0;
            //string vin4bu4zy4 = string.Concat(shang4gu3vin4bu4.Keys.AsEnumerable());
            
            foreach (var k in shang4gu3vin4bu4.Keys)//.Where(x=>x== "葉"))
            {
                string vin11 = k;
                string vin12 = k;
                Thread.Sleep(5000);
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + vin11 + "&y=" + vin12 + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, wbForSave, sheetNr, length, vin11, vin12);

                //ProcessTable("<table><tr><th><b style=\"戾<b style=\"戾", ws, wbForSave, sheetNr, length, vin11, vin12); 
                wbForSave.Worksheets[sheetNr].Name= vin11;
                wbForSave.Worksheets.Add();
                sheetNr++;
            }

            foreach (var tong1vin4 in tong1ia5)//.Where(x=>x== "魚屋"))
            {
                Thread.Sleep(5000);
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + tong1vin4[0] + "&y=" + tong1vin4[1] + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, wbForSave, sheetNr, length, tong1vin4[0].ToString(), tong1vin4[1].ToString());
                wbForSave.Worksheets[sheetNr].Name = tong1vin4;
                wbForSave.Worksheets.Add();
                sheetNr++;
            }

            string fn = shang4gu3vin4bu4.Keys.Count > 3 ? "shang4gu3vin4jo5" : string.Concat(shang4gu3vin4bu4.Keys.AsEnumerable());
            wbForSave.Save(@"D:\" + fn + "(" + DateTime.Now.ToString("yyMMddHHmm") + ").xlsx");
            ws.Workbook.Save(@"D:\shang4gu3li3in1(" + DateTime.Now.ToString("yyMMddHHmm") + ").xlsx");
        }

        private static Dictionary<string, int[]> GetPhoneticComponent(Worksheet ws, int exelRowsCount = 10000) //所有聲旁及占據的行
        {
            Dictionary<string, int[]> res = new Dictionary<string, int[]>();
            for (int j = 3; j < exelRowsCount; j++)
            {
                if (ws.Cells["A" + j.ToString()].Value != null)//"A"列是聲旁
                {
                    var shen1pang2 = ws.Cells["A" + j.ToString()].Value.ToString();
                    if (!string.IsNullOrEmpty(shen1pang2) && !res.ContainsKey(shen1pang2))
                    {
                        res.Add(shen1pang2, FindAllRowsOfSamePhoneticComponent(ws, j));
                    }
                    if (!shen1pang2vin4luei4.ContainsKey(shen1pang2))
                    {
                        List<string> onsets = new List<string>();
                        for (int i = res[shen1pang2][0]; i <= res[shen1pang2][1]; i++)
                        {
                            if (ws.Cells["G" + j.ToString()].Value != null)
                                onsets.Add(GetOnset(ws.Cells["G" + j.ToString()].Value.ToString()));
                        }
                        string shen1luei4 = GetOnsetGroup(onsets);
                        shen1pang2vin4luei4.Add(shen1pang2, shen1luei4);
                    }
                }
            }
            return res;
        }


        private static string GetOnsetGroup(string consonant)
        {
            if ("bpm".Any(x => consonant.Contains(x)))
            {
                return "P";
            }
            else if ("dtn".Any(x => consonant.Contains(x)))
            {
                return "T";
            }
            else if("gkŋɣh".Any(x=>consonant.Contains(x)))
            {
                return "K";
            }
            else if (consonant.Contains("l"))
            {
                return "L";
            }
            else if (consonant.Contains("x"))
            {
                return "K";
            }
            else if (consonant.Contains("s"))
            {
                return "S";
            }
            else if (consonant.Contains("r"))
            {
                return "R";
            }
            else
            {
                return "K";
            }
        }

        private static string GetOnsetGroup(List<string> onsets)//sg, rk, hr -> K
        {
            Dictionary<string,int> shen1luei4 = new Dictionary<string, int>() 
            {//聲類是聲母中的主輔音
                { "K", 0},
                { "P", 0},
                { "T", 0},
                { "L", 0},
                { "S", 0},
                { "R", 0},
            };

            double len = (double)onsets.Count;
            foreach (var onset in onsets)
            {
                var k = GetOnsetGroup(onset);
                shen1luei4[k]++;
                if ((double)((double)shen1luei4[k] / len) > 0.5)
                {
                    return k;
                }
            }
            return "";
        }

        private static void ProcessTable(string theText, Worksheet ws, Workbook wbForSave, int sheetNr, int length, string vin11, string vin12)
        {
            try
            {
                bool iao4gw3 = vin11 == vin12;//都是“魚”才可能要改音以保證此字屬於魚部
                Worksheet ws2 = wbForSave.Worksheets[sheetNr];
                var dt2 = ws2.Cells.ExportDataTable(0, 0, 1600, 9);
                string[] lines = theText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                string table = lines.Where(x => x.StartsWith("<table><tr><th")).FirstOrDefault();
                if (table != null)
                {
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

                            int lie5 = 0;
                            if (line.Contains("<i>"))
                            {
                                int shr3 = line.IndexOf("<i>") + 3;
                                int uei3 = line.IndexOf("</i>");
                                ws2.Cells[hang2, lie5].Value = line.Substring(shr3, uei3 - shr3);
                                lie5++;
                            }
                            for (int i = 0; i < rythms.Length - 1; i++)
                            {
                                string zy = GetCharacter(rythms[i]);
                                //if ("戾捩綟唳㑦蜧䓞悷".Contains(zy))
                                //{
                                //    int x = 0;
                                //}
                                Console.Write(zy);
                                CalcTotalHanzy(zy, ref vin4jo5zy4, ref cy3bu4zy4su4);
                                vals.Add(zy);
                                Dictionary<string, string[]> do1in1 = new Dictionary<string, string[]>();
                                for (int j = 1; j < length; j++)
                                {
                                    if (ws.Cells["P" + j.ToString()].Value == null || ws.Cells["G" + j.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                                        continue;
                                    else if (ws.Cells["P" + j.ToString()].Value.ToString().Contains(zy) &&
                                        ws.Cells["P" + j.ToString()].GetStyle().Font.Color != System.Drawing.ColorTranslator.FromHtml("#ffffcc00"))//"P"列是同聲旁同音字
                                    {
                                        var shang4gu3du5in1 = ws.Cells["G" + j.ToString()].Value.ToString(); //"G"列是上古音
                                        string[] zhong1gu3du5in1 = [ws.Cells["N" + j.ToString()].Value.ToString(), j.ToString()]; //"N"列是中古音
                                        if (!do1in1.Keys.Contains(shang4gu3du5in1))
                                            do1in1.Add(shang4gu3du5in1, zhong1gu3du5in1);
                                        if (shang4gu3vin4bu4[vin11].All(d => !shang4gu3du5in1.Contains(d))
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
                                if (iao4gw3 && do1in1.All(x => shang4gu3vin4bu4[vin11].All(d => !x.Key.Contains(d))))
                                {//多音字所有音都不合韻部，先嘗試人工智能修正，正不了再確定出韻
                                    bool i3siou1zhen4 = false;
                                    foreach (var gu3in1 in do1in1.ToList())
                                    {
                                        string qi2ta1shang4gu3in1 = FindRightOldPronunciation(ws, length, shang4gu3vin4bu4[vin11], gu3in1.Value, gu3in1.Key, length);
                                        if (qi2ta1shang4gu3in1 != gu3in1.Key && !do1in1.Keys.Contains(qi2ta1shang4gu3in1))
                                        {
                                            do1in1.Add(qi2ta1shang4gu3in1, gu3in1.Value);
                                            do1in1.Remove(gu3in1.Key);
                                            vals[vals.IndexOf(gu3in1.Key + "謬")] = qi2ta1shang4gu3in1;
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
                    ws2.Cells[hang2, 0].Value = iao4gw3 ? vin12 + "部紅色出韻音" + miou4su4.ToString() + "個" : "";
                    ws2.Cells[hang2, 8].Value = iao4gw3 ? chu5vin4zy4tong3ji4 : "";
                    ws2.Cells[hang2 + 1, 0].Value = vin11 + vin12 + "押韻韻腳字" + cy3bu4zy4su4.ToString() + "個: " + vin4jo5zy4;
                    Console.WriteLine(iao4gw3 ? chu5vin4zy4tong3ji4 : vin11 + vin12 + "通押如上");
                }
            }
            catch (Exception e)
            {
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
        private static string FindRightOldPronunciation(Worksheet ws, int length, string[] vin4bu4, string[] zhong1gu3in1, string shang4gu3in1, int exelRowsCount = 10000)
        {
            string res = shang4gu3in1;
            bool found = false;
            try
            {
                for (int j = 1; j < length; j++)
                {
                    if (ws.Cells["N" + j.ToString()].Value != null && //"N"列是中古音
                        ws.Cells["G" + j.ToString()].Value != null && //"G"列是上古音
                        ws.Cells["N" + j.ToString()].Value.Equals(zhong1gu3in1[0]) &&
                        vin4bu4.Any(x => ws.Cells["G" + j.ToString()].Value.ToString().Contains(x))
                        )
                    {//第一次自改：找同上古韻部的中古同音字
                        res = ws.Cells["G" + j.ToString()].Value.ToString();
                        ws.Cells["G" + zhong1gu3in1[1].ToString()].Value = res;
                        return res;
                    }
                }

                var zhong1gu3vin4mu3 = GetRhymeOfMC(zhong1gu3in1[0]);
                for (int j = 1; j < length; j++)
                {
                    if (ws.Cells["N" + j.ToString()].Value != null && //"N"列是中古音
                        ws.Cells["G" + j.ToString()].Value != null && //"G"列是上古音
                        ws.Cells["N" + j.ToString()].Value.ToString().Contains(zhong1gu3vin4mu3) &&
                        vin4bu4.Any(x => ws.Cells["G" + j.ToString()].Value.ToString().Contains(x))
                        )
                    {//第二次自改：找同上古韻部的中古同韻字
                        string vin4 = vin4bu4.First(x => ws.Cells["G" + j.ToString()].Value.ToString().Contains(x));
                        res = ChangePronuciationOC(ws, [vin4], shang4gu3in1, Convert.ToInt32(zhong1gu3in1[1]), ref found, exelRowsCount);
                        if (found)
                            return res;
                    }
                }

                //第三次自改： 無據試改
                res = ChangePronuciationOC(ws, vin4bu4, shang4gu3in1, Convert.ToInt32(zhong1gu3in1[1]), ref found, exelRowsCount);
                return res;
            }
            catch (Exception e)
            {
                return res;
            }

        }

        private static string ChangePronuciationOC(Worksheet ws, string[] vin4bu4, string shang4gu3in1, int rowWithZy, ref bool found, int exelRowsCount = 10000)
        {
            string res = shang4gu3in1;
            foreach (var v in shang4gu3vin4bu4)
            {
                foreach (var rhyme in v.Value)
                {
                    if (shang4gu3in1.Contains(rhyme))
                    {
                        foreach (var vin4 in vin4bu4)
                        {
                            if ((shang4gu3vin4bu4["覺"].Contains(rhyme) || shang4gu3vin4bu4["藥"].Contains(rhyme)) &&
                                (shang4gu3vin4bu4["覺"].Contains(vin4) || shang4gu3vin4bu4["藥"].Contains(vin4)) &&
                                "錫覺".Contains(ws.Cells["K" + rowWithZy.ToString()].Value.ToString()))
                            {
                                continue;
                            }
                            if ((shang4gu3vin4bu4["鐸"].Contains(rhyme) || shang4gu3vin4bu4["藥"].Contains(rhyme)) &&
                                (shang4gu3vin4bu4["鐸"].Contains(vin4) || shang4gu3vin4bu4["藥"].Contains(vin4)) &&
                                "鐸".Contains(ws.Cells["K" + rowWithZy.ToString()].Value.ToString()))
                            {//鑿不可改成skˤøk，ˤøk是四開錫，韻部歸納有誤
                                continue;
                            }
                            if ((shang4gu3vin4bu4["鐸"].Contains(rhyme) || shang4gu3vin4bu4["藥"].Contains(rhyme)) &&
                                (shang4gu3vin4bu4["鐸"].Contains(vin4) || shang4gu3vin4bu4["藥"].Contains(vin4)) &&
                                "鐸".Contains(ws.Cells["K" + rowWithZy.ToString()].Value.ToString()))
                            {//鑿不可改成skˤøk，ˤøk是四開錫，韻部歸納有誤
                                continue;
                            }
                            var sin1in1 = shang4gu3in1.Replace(rhyme, vin4);
                            if (RythmsOC().Any(x => sin1in1.EndsWith(x)))
                            {
                                ws.Cells["G" + rowWithZy.ToString()].Value = res = sin1in1;
                                if (!DoubleMapping(ws, res, exelRowsCount))         //防止一上古對多中古
                                {
                                    found = true;
                                    var zy4zu5hang2 = FindAllRowsOfSamePhoneticComponent(ws, rowWithZy);
                                    ChangePronuciationOfBrothers(ws, zy4zu5hang2, rhyme, vin4, exelRowsCount);
                                    return res;
                                }
                                else
                                {
                                    ws.Cells["G" + rowWithZy.ToString()].Value = res = shang4gu3in1;
                                }
                            }
                        }
                    }
                }
            }
            return res;
        }

        static void ChangePronuciationOfBrothers(Worksheet ws, int[] rows, string rhyme, string newryhme, int exelRowsCount = 10000)
        {
            int hang2 = rows[0];
            while (hang2 <= rows[1])
            {
                string sin1in1zie5 = ws.Cells["G" + hang2.ToString()].Value.ToString().Replace(rhyme, newryhme);
                if (RythmsOC().Any(x => sin1in1zie5.EndsWith(x)))
                {
                    if ((ws.Cells["L" + hang2.ToString()].ToString().Equals("入") && "ptk".Any(x => sin1in1zie5.EndsWith(x))) ||
                        (ws.Cells["L" + hang2.ToString()].ToString().Equals("平") && "ptkhɣ".All(x => !sin1in1zie5.EndsWith(x))) ||
                        (ws.Cells["L" + hang2.ToString()].ToString().Equals("上") && "ptkh".All(x => !sin1in1zie5.EndsWith(x))) ||
                        ws.Cells["L" + hang2.ToString()].ToString().Equals("去"))
                    {
                        ws.Cells["G" + hang2.ToString()].Value = sin1in1zie5;
                    }
                }
                if (DoubleMapping(ws, sin1in1zie5, exelRowsCount))
                    ws.Cells["G" + hang2.ToString()].Value = sin1in1zie5.Replace(newryhme, rhyme);
                hang2++;
            }
        }


        private static int[] FindAllRowsOfSamePhoneticComponent(Worksheet ws, int row)
        {
            var res = new int[2];
            var gw5 = ws.Cells["A" + row.ToString()];
            if (gw5.IsMerged)
            {//如：缶族9-17行
                var gw5sin4si5 = gw5.GetMergedRange();
                res[0] = gw5sin4si5.FirstRow + 1;
                res[1] = gw5sin4si5.FirstRow + gw5sin4si5.RowCount;
            }
            else
            {//如：孑族只有一行krat見三開薛B入居列kiät孑𨥂
                res[0] = res[1] = row;
            }
            return res;
        }

        private static List<string> RythmsOC() //所有上古韻母
        {
            var res = new List<string>();
            foreach (var vin4bu in shang4gu3vin4bu4.Values)
            {
                foreach (var v in vin4bu)
                {
                    res.Add(v);
                    res.Add(v + "h");
                    if ("ktp".All(x => !v.EndsWith(x)))
                        res.Add(v + "ɣ");
                }
            }
            return res;
        }
        
        private static Dictionary<string,int> OnsetsOC(Worksheet ws, int exelRowsCount = 10000) //所有當前擬構的上古聲母及出現次數
        {
            Dictionary<string, int> res = new Dictionary<string, int>();

            for (int j = 3; j < exelRowsCount; j++)
            {
                if (ws.Cells["G" + j.ToString()].Value != null)//"G"列是上古音
                {
                    var in1zie5 = ws.Cells["G" + j.ToString()].Value.ToString();
                    string shen1 = GetOnset(in1zie5);
                    if (!res.Keys.Contains(shen1))
                        res.Add(shen1, 1);
                    else
                        res[shen1]++;
                }
            }
            return res;
        }

        static string GetOnset(string in1zie5)//sgat->sg
        {
            List<string> vin4mu3lie5bao3 = new List<string>();
            foreach (var item in RythmsOC())
            {
                vin4mu3lie5bao3.Add("ˤ" + item);
                vin4mu3lie5bao3.Add(item);
            }
            string res = "";
            foreach (var v in vin4mu3lie5bao3)
            {
                if (in1zie5.IndexOf(v) > 0)
                {
                    res = in1zie5.Substring(0, in1zie5.IndexOf(v));
                }
            }
            return res;
        }

        private static string GetRhymeOfMC(string syllable)
        {
            string syl = RemoveDiacritics(syllable);
            int idx = syl.Length - 1;
            foreach (var c in zhong1gu3vwn2in1)
            {
                if (syl.IndexOf(c) > -1)
                    idx = Math.Min(idx, syl.IndexOf(c));
            }
            return syl.Substring(idx);

        }

        private static bool DoubleMapping(Worksheet ws, string shang4gu3in1, int length = 10000)
        {
            bool res = false;
            Dictionary<string, List<string>> Mapping = new Dictionary<string, List<string>>() { { shang4gu3in1, new List<string>() } };
            for (int j = 1; j < length; j++)
            {
                if (ws.Cells["P" + j.ToString()].Value == null || ws.Cells["G" + j.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                    continue;

                if (ws.Cells["G" + j.ToString()].Value != null && ws.Cells["G" + j.ToString()].Value.ToString() == shang4gu3in1)//"G"列是上古音
                {
                    string k = ws.Cells["G" + j.ToString()].Value.ToString();//"G"列是上古音
                    if (ws.Cells["N" + j.ToString()].Value != null)//"N"列是中古音
                    {
                        string v = ws.Cells["N" + j.ToString()].Value.ToString() + ws.Cells["K" + j.ToString()].Value.ToString();//"K"列是中古韻
                        //if (v.Contains("lvì"))
                        //    rowNo = 1;
                        //if (v.Contains("lièi"))
                        //    rowNo = 2;

                        if (!Mapping[shang4gu3in1].Contains(v))
                        {
                            Mapping[shang4gu3in1].Add(v);
                        }
                    }
                }
            }
            foreach (var kv in Mapping)
            {
                if (kv.Value.Count > 1)
                {
                    res = true;
                }
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
                        string v = ws.Cells["N" + res.ToString()].Value.ToString() + ws.Cells["K" + res.ToString()].Value.ToString();//"K"列是中古韻
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
                if (kv.Value.Count > 1)
                {
                    Console.WriteLine("以下的上古擬音對映多個中古音，請重擬！");
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

        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder(capacity: normalizedString.Length);

            for (int i = 0; i < normalizedString.Length; i++)
            {
                char c = normalizedString[i];
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark || c == '̈')
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder
                .ToString()
                .Normalize(NormalizationForm.FormC);
        }

        static string GetCharacter(string text)
        {

            string zy = text.Substring(text.Length - 1);
            if (zy == "A" || zy == "B")
                zy = text.Substring(text.Length - 2, 1);
            if (zy.Contains("\ude62") || zy.Contains("\udfae"))
            {
                zy = text.Substring(text.Length - 2, 2);
            }
            if (zy.Contains("}"))
            {
                zy = text.Substring(text.Length - 3, 3);
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
            else if (zy.Contains("糸費"))
            {
                zy = "紼";
            }
            else if (zy == "卝")
            {
                zy = "丱";
            }
            return zy;
        }

        static void shang4gu3duei4zhong1gu3(Worksheet ws, int length)
        {

            Workbook wbForSave = new Workbook();
            wbForSave.Worksheets.Add();
            //var ws0 = wbForSave.Worksheets[0];
            //var ws1 = wbForSave.Worksheets[1];
            List<string> sheng1mu3=new List<string>();

            int res = 3;
            int nullcount = 0;
            while ((ws.Cells["G" + res.ToString()].Value == null ||
                       !String.IsNullOrWhiteSpace(ws.Cells["G" + res.ToString()].Value.ToString()))&& res<=length) //"G"列是上古音
            {
                if (ws.Cells["G" + res.ToString()].Value != null)//"G"列是上古音
                {
                    string k = ws.Cells["G" + res.ToString()].Value.ToString();//"G"列是上古音
                    if ((k.EndsWith("h") && !k.EndsWith("kh") && !k.EndsWith("th") && !k.EndsWith("ph")) || k.EndsWith("ɣ"))
                        k = k.Substring(0, k.Length - 1);

                    bool shenmuFound = false;
                    foreach (var kv in shen1mu3duei4in4)
                    {
                        foreach (var shang4gu3sheng1mu3 in kv.Value)
                        {
                            if (k.StartsWith(shang4gu3sheng1mu3))
                            {
                                shenmuFound = true;
                                //string cellvalue = "";
                                bool vinmuFound = false;
                                foreach (var vin5bu4 in shang4gu3vin4bu4)
                                {
                                    foreach(var v in vin5bu4.Value)
                                    {
                                        if (k.EndsWith(v))
                                        {
                                            int hang2hao4=0;
                                            vinmuFound=true;
                                            var vinmu = k.Substring(shang4gu3sheng1mu3.Length);//e.g.: am, ʷˤəl ...
                                            foreach (var vinbu in vin4bu4hao4)
                                            {
                                                if (vinmu.EndsWith(vinbu.Key))
                                                {
                                                    hang2hao4 = vinbu.Value;
                                                    break;
                                                }
                                            }
                                            string ss = vinmu + ws.Cells["I" + res.ToString()].Value.ToString()
                                                + ws.Cells["J" + res.ToString()].Value.ToString() + ws.Cells["K" + res.ToString()].Value.ToString();
                                            int biao3hao4 = k.Contains("ˤ") ? 1 : 0;
                                            if (wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()].Value == null||
                                                !wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()].Value.ToString().Contains(ss))
                                                wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()].Value += ss + ", ";
                                            break;
                                        }
                                    }
                                    if (vinmuFound)
                                        break;
                                }
                                break;
                            }
                        }
                        if (shenmuFound)
                            break;

                    }
                }
                res++;
            }
            wbForSave.Save(@"D:\上古對中古" + DateTime.Now.ToString("yyMMddHHmm") + ".xlsx");
        }
    }
}
