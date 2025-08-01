﻿using Shang4Gu3In1;
using Aspose.Cells;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace Shang4Gu3In1
{
    class Program
    {


        const string lu5vwn2in1 = "aɛɔəøo";//六元音
        private static Dictionary<string, string> liou3mang4vin4bu4 = new Dictionary<string, string>() {
              { "鐸", "ak"}, { "錫", "ɛk"}, { "屋", "ɔk"}, { "職", "ək"}, { "藥", "øk"}, { "覺", "ok"},
              { "陽", "aŋ"}, { "耕", "ɛŋ"}, { "東", "ɔŋ"}, { "蒸", "əŋ"}, { "冬", "oŋ"},
              { "歌(歌0)", "al"}, { "皮(歌1)","ɛl"},  { "微", "əl"}, { "脂", "øl"},
              { "月(月0)", "at"}, { "拜(月1)","ɛt"}, { "物", "ət"}, { "質", "øt"},
              { "元(元0)", "an"}, { "見(元1)","ɛn"}, { "文", "ən"},
              { "真(真0)", "øn"}, { "印(真1)","øŋ"},
              { "葉(葉0)", "ap"}, { "業(葉1)","ɛp"}, { "緝", "əp"},
              { "談(談0)", "am"}, { "嚴(談1)","ɛm"}, { "侵", "əm"},
              { "魚", "a"}, { "之", "ə"}, { "支", "ɛ"},
              { "侯", "ɔ"}, { "幽", "o"}, { "宵", "ø"},
        };

        private static Dictionary<string, string[]> shang4gu3vin4bu4 = new Dictionary<string, string[]>() {//三十一韻部
             { "鐸", [lu5vwn2in1[0]+"k"]}, { "錫", [lu5vwn2in1[1]+"k"]}, { "屋", [lu5vwn2in1[2]+"k"]}, { "職", [lu5vwn2in1[3]+"k"]}, { "藥", [lu5vwn2in1[4]+"k"]}, { "覺", [lu5vwn2in1[5]+"k"]},
             { "陽", [lu5vwn2in1[0]+"ŋ"]}, { "耕", [lu5vwn2in1[1]+"ŋ"]}, { "東", [lu5vwn2in1[2]+"ŋ"]}, { "蒸", [lu5vwn2in1[3]+"ŋ"]}, { "冬", [lu5vwn2in1[5]+"ŋ"]},
             { "歌", [lu5vwn2in1[0]+"l", lu5vwn2in1[1]+"l"]},  { "微", [lu5vwn2in1[3]+"l"]}, { "脂", [lu5vwn2in1[4]+"l"]},
             { "月", [lu5vwn2in1[0]+"t", lu5vwn2in1[1]+"t"]}, { "物", [lu5vwn2in1[3]+"t"]}, { "質", [lu5vwn2in1[4]+"t"]},
             { "元", [lu5vwn2in1[0]+"n", lu5vwn2in1[1]+"n"]}, { "文", [lu5vwn2in1[3]+"n"]},
             { "真", [lu5vwn2in1[4]+"n", lu5vwn2in1[4]+"ŋ"]},
             { "葉", [lu5vwn2in1[0]+"p", lu5vwn2in1[1]+"p"]}, { "緝", [lu5vwn2in1[3]+"p"]},
             { "談", [lu5vwn2in1[0]+"m", lu5vwn2in1[1]+"m"]}, { "侵", [lu5vwn2in1[3]+"m"]},
             { "魚", [lu5vwn2in1[0].ToString()]}, { "之", [lu5vwn2in1[3].ToString()]}, { "支", [lu5vwn2in1[1].ToString()]},
             { "侯", [lu5vwn2in1[2].ToString()]}, { "幽", [lu5vwn2in1[5].ToString()]}, { "宵", [lu5vwn2in1[4].ToString()]},
        };

        private static Dictionary<string, int> vin4bu4hao4 = new Dictionary<string, int>() {//三十六韻
             {lu5vwn2in1[0]+"k", 1}, { lu5vwn2in1[1]+"k", 2}, { lu5vwn2in1[2]+"k",3}, { lu5vwn2in1[3]+"k", 4}, { lu5vwn2in1[4]+"k",5}, { lu5vwn2in1[5]+"k",6},
             { lu5vwn2in1[0]+"ŋ", 7}, { lu5vwn2in1[1]+"ŋ", 8}, { lu5vwn2in1[2]+"ŋ", 9}, { lu5vwn2in1[3]+"ŋ", 10}, { lu5vwn2in1[5]+"ŋ", 11},
             { lu5vwn2in1[0]+"l",12 }, {lu5vwn2in1[1]+"l",13},  { lu5vwn2in1[3]+"l",14}, { lu5vwn2in1[4]+"l",15},
             { lu5vwn2in1[0]+"t",16 },{ lu5vwn2in1[1]+"t",17}, { lu5vwn2in1[3]+"t",18}, { lu5vwn2in1[4]+"t", 19},
             { lu5vwn2in1[0]+"n", 20 },{ lu5vwn2in1[1]+"n", 21 }, { lu5vwn2in1[3]+"n",22},
             { lu5vwn2in1[4]+"n", 23 },{ lu5vwn2in1[4]+"ŋ", 24},
             { lu5vwn2in1[0]+"p", 25 }, { lu5vwn2in1[1]+"p",26}, {lu5vwn2in1[3]+"p",27},
             { lu5vwn2in1[0]+"m", 28}, { lu5vwn2in1[1]+"m", 29 },{lu5vwn2in1[3]+"m", 30 },
             { lu5vwn2in1[0].ToString(), 31}, { lu5vwn2in1[3].ToString(), 32}, { lu5vwn2in1[1].ToString(), 33},
             { lu5vwn2in1[2].ToString(), 34}, { lu5vwn2in1[5].ToString(), 35}, { lu5vwn2in1[4].ToString(), 36},
        };

        private static Dictionary<string, string[]> shen1mu3duei4in4 = new Dictionary<string, string[]>() {//上古中古聲母對映
            { "SKR莊組A", ["skʀ", "skʰʀ", "sgʀ", "sxʀ"]},
            { "SK精組B", ["skʰ", "sk", "sg"]},
            { "SK精組C", ["sx", "sɣ"]},
            { "TR知組D", ["tʀ", "tʰʀ", "dʀ", "nʀ"]},
            { "SR生母E", ["sʀ"]},
            { "KL章端組母F", ["kl", "kʰl", "gl", "ŋl"]},
            { "XL透書船母G", ["xl", "xn", "hl"]},
            { "ƔL定以母H", [ "ɣl", "l"]},
            { "T章端組I", [ "tʰ", "t", "d"]},
            { "N明泥日娘組J", ["n"]},
            { "R來母K", ["ʀ"]},
            { "S心生L", ["s"]},
            { "P幫組M", ["pʰ", "p", "b"]},
            { "M明母N", ["m"]},
            { "h影母O", ["h"]},
            { "X曉P", ["x"]},
            { "Ɣ匣云Q", ["ɣ"]},
            { "K見溪群R", ["kʰ", "k",  "g"]},
            { "Ŋ疑明S", ["ŋ"]},
        };

        static List<string> tong1ia5 = new List<string>() { "魚鐸", "魚陽", "魚之", "魚支", "魚侯", "魚屋", "魚東", "魚幽", "魚宵", "魚歌", "魚元", "魚微",
            "鐸之", "鐸職", "鐸錫", "鐸侯", "鐸屋", "鐸幽", "鐸藥", "鐸歌", "鐸質", "鐸葉", "陽蒸", "陽錫", "陽耕", "陽東", "陽冬", "陽真",
            "之職", "之蒸", "之侯", "之幽", "之覺", "之宵", "之元", "之緝", "職蒸", "職侯", "職屋", "職幽", "職覺", "職藥", "職葉", "職緝", 
            "蒸侯", "蒸東", "蒸冬", "蒸文", "蒸侵", "支錫", "支歌", "支月", "支元", "支微", "支物", "支文", "支脂", "支質",
            "錫屋", "錫歌", "錫月", "錫物", "錫質", "耕東", "耕元", "耕文", "耕真", "侯屋", "侯東", "侯幽", "侯冬", "侯宵", 
            "屋覺", "屋宵", "屋藥", "東幽", "東冬", "東侵", "幽覺", "幽宵", "覺宵", "覺緝",  "冬侵", "宵元", "歌元", "歌微", "歌物", "歌脂", 
            "月元", "月物", "月脂", "月質", "月葉", "月緝", "元微", "元物", "元文", "元脂", "元質", "元真", "元緝", "微物", "微文", "微脂", "微質", 
            "物脂", "物質", "文質", "文真", "文緝", "脂質",  "葉談", "葉緝", "談侵"};
        //kaom通押判斷錯誤，刪 "冬真", "質真",

        private static string zhong1gu3vwn2in1 = "aeiouvwryäüöëï";//廣通中古拼音的元音

        private static Dictionary<string, string> shen1pang2vin4luei4 = new Dictionary<string, string>();

        static async Task Main(string[] args)
        {
#pragma region 按聲旁筆畫數排序
           /* var myDict = ReadCsvToDictionary("D:/output.csv").OrderBy(x=>x.Value);
            Console.OutputEncoding = Encoding.UTF8;
            Workbook wk = new Workbook("D:/廣韻字上古音形考.xlsx");
            Worksheet ws = wk.Worksheets[0];
            MoveRedCharactersToFrontInColumn(ws, 15);
;           //Din4Vin4Bu4(ws, 4, 6);
            //UnmergeAndPropagateValueInColumn(ws, 1);
            /int startRow = 2;
            foreach (var kv in myDict)//.Where(x=>x.Value>2))
            {
                var (rowIndex, rowCount) = FindRowAndMergedLengthByPrefix(ws, kv.Key, 0);
                CutAndInsertRows(ws, rowIndex, rowCount, ref startRow);
                startRow += rowCount;
            }*
            wk.Save("D:/廣韻字上古音形考1.xlsx");*/
#pragma endregion 按聲旁筆畫數排序
            Console.OutputEncoding = Encoding.UTF8;
            Workbook wk = new Workbook("D:/廣韻字上古音形考.xlsx");
            Worksheet ws = wk.Worksheets[0];
            //CheckDen(ws);
            int length = CheckDoubleMapping(ws);
            var d = OnsetsOC(ws, length);
            foreach (var s in d.OrderBy(x => x.Key).ThenBy(x => x.Value.Sum(d => d.Value)))
            {
                Console.WriteLine(s.Key + ":" + s.Value.Sum(d => d.Value));
                foreach (var kv in s.Value.OrderBy(v => v.Value))
                {
                    Console.WriteLine(kv);
                }
            }
            /* foreach (var s in GetPhoneticComponent(ws,length))
            {
                Console.WriteLine(s);
             }*/
            /*
            foreach (var s in shen1pang2vin4luei4)
            {
                 Console.WriteLine(s);
            }*/

            var shenmuZhongDueiShang = ShengMuZhongDueiShang(ws, length);
            foreach (var dd in shenmuZhongDueiShang.OrderBy(x=>x.Key))
            {
                Console.WriteLine(dd.Key);
                foreach (var ee in dd.Value)
                {
                    Console.WriteLine(ee.Key + ":" + ee.Value);
                }
            }
                //Console.WriteLine(lu5vwn2in1[1]);

            //Huang4Üin4(new List<string>() { "三開嚴" }, "əm", "øm");
            var vinbu2denvin = shang4gu3duei4zhong1gu3(ws, length);
            int sheetNr = 0;
            //string vin4bu4zy4 = string.Concat(shang4gu3vin4bu4.Keys.AsEnumerable());

            Workbook wbForSave = new Workbook();
            foreach (var k in shang4gu3vin4bu4.Keys.Where(x=>x=="真"))
            {
                string vin11 = k;
                string vin12 = k;
                Thread.Sleep(5000);
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + vin11 + "&y=" + vin12 + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, wbForSave, vinbu2denvin, ref sheetNr, length, vin11, vin12, false);

                //ProcessTable("<table><tr><th><b style=\"戾<b style=\"戾", ws, wbForSave, sheetNr, length, vin11, vin12); 
            }
            
            foreach (var tong1vin4 in tong1ia5.Where(x=>x.Contains("真")))
            {
                Thread.Sleep(5000);                
                var httpResponseMessage = await DataService.Client.GetAsync("http://www.kaom.net/yayuns_bu88.php?book=all&x=" + tong1vin4[0] + "&y=" + tong1vin4[1] + "&mode=yunbu");
                var content = await httpResponseMessage.Content.ReadAsStringAsync();
                ProcessTable(content, ws, wbForSave, vinbu2denvin, ref sheetNr, length, tong1vin4[0].ToString(), tong1vin4[1].ToString(), false);
            }

            string fn = shang4gu3vin4bu4.Keys.Count > 3 ? "上古韻腳" : string.Concat(shang4gu3vin4bu4.Keys.AsEnumerable());
            wbForSave.Save(@"D:\" + fn  + DateTime.Now.ToString("yyMMddHHmm") + ".xlsx");
            //ws.Workbook.Save(@"D:\shang4gu3li3in1(" + DateTime.Now.ToString("yyMMddHHmm") + ").xlsx");
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
            else if (consonant.Contains("ʀ"))
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

        private static void ProcessTable(string theText, Worksheet cvwn2zy4biao3, Workbook wbForVinJo, Workbook vinbu2denvin, ref int sheetNr, int length, string vin11, string vin12,bool aiModificating = true)
        {
            try
            {
                int startSheetNr = sheetNr;
                int miou4su4 = 0;
                int chu5vin4zy4su4 = 0, cy3bu4zy4su4 = 0;
                string chu5vin4zy4 = "", vin4jo5zy4 = "";
                bool iao4gw3 = vin11 == vin12;//都是“魚”才可能"要改"音以保證此字屬於魚部
                string[] lines = theText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                string table = lines.Where(x => x.StartsWith("<table><tr><th")).FirstOrDefault();
                if (!iao4gw3 || shang4gu3vin4bu4[vin11].Length == 1)
                {
                    if (table != null)
                    {
                        Worksheet wsForVinJo = wbForVinJo.Worksheets[sheetNr];
                        lines = table.Split(new string[] { "<tr><td>" }, StringSplitOptions.None);
                        int hang2 = 0;
                        foreach (string line in lines)
                        {
                            var rythms = line.Split(new string[] { "<b style=\"" }, StringSplitOptions.None);
                            List<string> vals = new List<string>();
                            if (rythms.Length > 1)
                            {

                                int lie5 = FindColumnSetValue(wsForVinJo, line, hang2);
                                
                                for (int i = 0; i < rythms.Length - 1; i++)
                                {
                                    string zy = GetCharacter(rythms[i]);
                                    Dictionary<string, string[]> do1in1 = Chu3Li3Vin4Jo5Zy4(cvwn2zy4biao3, zy, i, length, vals, vin11, ref miou4su4, ref vin4jo5zy4, ref cy3bu4zy4su4);                                   
                                    if (iao4gw3 && do1in1.All(x => shang4gu3vin4bu4[vin11].All(d => !x.Key.Contains(d))))
                                    {//多音字所有音都不合韻部，先嘗試人工智能修正，正不了再確定出韻
                                        bool i3siou1zhen4 = false;
                                        foreach (var gu3in1 in do1in1.ToList())
                                        {
                                            string qi2ta1shang4gu3in1 = aiModificating ? FindRightOldPronunciation(cvwn2zy4biao3, vinbu2denvin, length, shang4gu3vin4bu4[vin11], gu3in1.Value, gu3in1.Key, length) : gu3in1.Key;
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
                                hang2 = Miou4Hong2(vals, wsForVinJo, hang2, ref lie5);
                            }
                        }
                        Tong3Ji4Chu5Vin4Zy4(vin11, vin12, chu5vin4zy4su4, chu5vin4zy4, wsForVinJo, hang2, iao4gw3, miou4su4, cy3bu4zy4su4, vin4jo5zy4);
                    }
                    AddSheet(wbForVinJo, vin11 + vin12, ref sheetNr);
                }
                else if (shang4gu3vin4bu4[vin11].Length > 1)//韻部細分
                {
                    if (table != null)
                    {
                        lines = table.Split(new string[] { "<tr><td>" }, StringSplitOptions.None);
                        List<int> hang2 = new List<int>();
                        int sinieshu = shang4gu3vin4bu4[vin11].Length;// + CalculateCombination(shang4gu3vin4bu4[vin11].Length, 2);//歌：al，El，al+El 3新頁
                        for (int i = 0; i <= sinieshu; i++)
                        {
                            AddSheet(wbForVinJo, i!=2 ? vin11 + i : vin11 + "混", ref sheetNr);
                            hang2.Add(0);
                        }

                        foreach (string line in lines)
                        {
                            var rythms = line.Split(new string[] { "<b style=\"" }, StringSplitOptions.None);
                            List<string> vals = new List<string>();
                            List<List<string>> duin = new List<List<string>>();
                            if (rythms.Length > 1)
                            {
                               
                                for (int i = 0; i < rythms.Length - 1; i++)
                                {//處理一首詩所有韻腳字
                                    string zy = GetCharacter(rythms[i]);
                                    List<string> guinmen = new List<string>();
                                    Dictionary<string, string[]> do1in1 = Chu3Li3Vin4Jo5Zy4(cvwn2zy4biao3, zy, i, length, vals, vin11, ref miou4su4, ref vin4jo5zy4, ref cy3bu4zy4su4, guinmen);                                    
                                    duin.Add(guinmen);
                                }
                                bool u2tong1ia5 = false;
                                for (int jj = 0; jj < shang4gu3vin4bu4[vin11].Length; jj++)
                                {

                                    if (duin.All(x => x.Any(ele => ele.Contains(shang4gu3vin4bu4[vin11][jj]))))
                                    {
                                        u2tong1ia5 = true;
                                        var wsForVinJo = wbForVinJo.Worksheets[startSheetNr + jj];
                                        int lie5 = FindColumnSetValue(wsForVinJo, line, hang2[jj]);
                                        hang2[jj] = Miou4Hong2(vals, wsForVinJo, hang2[jj], ref lie5);
                                        break;
                                    }
                                }
                                if (!u2tong1ia5)
                                {
                                    int idx = hang2.Count - 1;
                                    var wsForVinJo = wbForVinJo.Worksheets[startSheetNr + idx];
                                    int lie5 = FindColumnSetValue(wsForVinJo,line, hang2[idx]);
                                    if (duin.Where(x => x.Count > 0).Any(y => y.All(x => x.Contains("謬"))))
                                    {//多音字所有音都不合韻部，先嘗試人工智能修正，正不了再確定出韻
                                        var itms = duin.Where(x => x.Count > 0 && x.All(s => s.Contains("謬"))).ToList();
                                        foreach (var itm in itms)
                                        {
                                            int svhao = duin.IndexOf(itm);
                                            var zy = rythms[svhao].Substring(rythms[svhao].Length - 1);
                                            CalcTotalHanzy(zy, ref chu5vin4zy4, ref chu5vin4zy4su4);
                                            vals[vals.IndexOf(zy)] += "謬";                                            
                                        }
                                    }
                                    hang2[idx] = Miou4Hong2(vals, wsForVinJo, hang2[idx], ref lie5);
                                }
                                Console.WriteLine();
                                
                            }
                        }
                        Tong3Ji4Chu5Vin4Zy4(vin11, vin12, chu5vin4zy4su4, chu5vin4zy4, wbForVinJo.Worksheets[startSheetNr + hang2.Count - 1], hang2[hang2.Count - 1], iao4gw3, miou4su4, cy3bu4zy4su4, vin4jo5zy4);
                    }
                }
            }
            catch (Exception e)
            {
            }
        }

        /*
         * 統計出韻字
         */
        private static void Tong3Ji4Chu5Vin4Zy4(string vin11, string vin12, int chu5vin4zy4su4, string chu5vin4zy4, Worksheet wsForVinJo, int hang2, bool iao4gw3, int miou4su4, int cy3bu4zy4su4, string vin4jo5zy4)
        {
            string chu5vin4zy4tong3ji4 = vin12 + "部紅色出韻字" + chu5vin4zy4su4.ToString() + "個：" + chu5vin4zy4;
            wsForVinJo.Cells[hang2, 0].Value = iao4gw3 ? vin12 + "部紅色出韻音" + miou4su4.ToString() + "個" : "";
            wsForVinJo.Cells[hang2, 8].Value = iao4gw3 ? chu5vin4zy4tong3ji4 : "";
            wsForVinJo.Cells[hang2 + 1, 0].Value = vin11 + vin12 + "押韻韻腳字" + cy3bu4zy4su4.ToString() + "個: " + vin4jo5zy4;
            Console.WriteLine(iao4gw3 ? chu5vin4zy4tong3ji4 : vin11 + vin12 + "通押如上");
        }

        /*
         * 謬紅         
         */
        private static int Miou4Hong2(List<string> vals, Worksheet wsForVinJo, int hang2, ref int lie5)
        {
            foreach (var v in vals)
            {
                if (v.EndsWith("謬"))
                {
                    Style style = new Style();
                    style.Font.Color = Color.Red;
                    wsForVinJo.Cells[hang2, lie5].SetStyle(style);
                }
                wsForVinJo.Cells[hang2, lie5].Value = v.Replace("謬", "");
                lie5++;
            }
            hang2++;
            return hang2;
        }

        private static void AddSheet(Workbook wbForVinJo, string sheetName, ref int sheetNr)
        {
            wbForVinJo.Worksheets[sheetNr].Name = sheetName;
            wbForVinJo.Worksheets.Add();
            sheetNr++;
        }

        private static int FindColumnSetValue(Worksheet wsForVinJo, string line, int column)
        {
            int lie5 = 0;
            if (line.Contains("<i>"))
            {
                int shr3 = line.IndexOf("<i>") + 3;
                int uei3 = line.IndexOf("</i>");
                wsForVinJo.Cells[column, lie5].Value = line.Substring(shr3, uei3 - shr3);
                lie5++;
            }
            return lie5;
        }
        /*
         * 處理韻腳字    
         */
        private static Dictionary<string, string[]> Chu3Li3Vin4Jo5Zy4(Worksheet cvwn2zy4biao3,string zy, int i, int length, List<string> vals, string vin11, ref int miou4su4, ref string vin4jo5zy4, ref int cy3bu4zy4su4, List<string>? guinmen = null)//處理一首詩所有韻腳字
        {
            Console.Write(zy);
            CalcTotalHanzy(zy, ref vin4jo5zy4, ref cy3bu4zy4su4);
            vals.Add(zy);            
            Dictionary<string, string[]> do1in1 = new Dictionary<string, string[]>();
            for (int j = 3; j < length; j++)
            {//處理了一字所有古音
                if (cvwn2zy4biao3.Cells["P" + j.ToString()].Value == null || cvwn2zy4biao3.Cells["G" + j.ToString()].Value == null)//"P"列是同聲旁同音字"G"列是上古音
                    continue;
                else if (cvwn2zy4biao3.Cells["P" + j.ToString()].Value.ToString().Contains(zy) &&
                    cvwn2zy4biao3.Cells["P" + j.ToString()].GetStyle().Font.Color != System.Drawing.ColorTranslator.FromHtml("#ffffcc00"))//"P"列是同聲旁同音字
                {
                    var shang4gu3du5in1 = cvwn2zy4biao3.Cells["G" + j.ToString()].Value.ToString(); //"G"列是上古音
                    string[] zhong1gu3du5in1 = [cvwn2zy4biao3.Cells["N" + j.ToString()].Value.ToString(), j.ToString()]; //"N"列是中古音
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
                    guinmen?.Add(shang4gu3du5in1);
                }
            }
            return do1in1;
        }

        private static void CalcTotalHanzy(string zy, ref string so3iou3zy4, ref int su4liang4)
        {
            if (!so3iou3zy4.Contains(zy))
            {
                su4liang4++;
                so3iou3zy4 += zy;
            }
        }
        private static string FindRightOldPronunciation(Worksheet cvwn2zy4biao3, Workbook vinbu2denvin, int length, string[] vin4bu4, string[] zhong1gu3in1, string shang4gu3in1, int exelRowsCount = 10000)
        {
            string res = shang4gu3in1;
            bool found = false;
            try
            {
                for (int j = 3; j < length; j++)
                {
                    if (cvwn2zy4biao3.Cells["N" + j.ToString()].Value != null && //"N"列是中古音
                        cvwn2zy4biao3.Cells["G" + j.ToString()].Value != null && //"G"列是上古音
                        cvwn2zy4biao3.Cells["N" + j.ToString()].Value.Equals(zhong1gu3in1[0]) &&
                        vin4bu4.Any(x => cvwn2zy4biao3.Cells["G" + j.ToString()].Value.ToString().Contains(x))
                        )
                    {//第一次自改：找同上古韻部的中古同音字
                        res = cvwn2zy4biao3.Cells["G" + j.ToString()].Value.ToString();
                        cvwn2zy4biao3.Cells["G" + zhong1gu3in1[1].ToString()].Value = res;
                        return res;
                    }
                }

                var zhong1gu3vin4mu3 = GetRhymeOfMC(zhong1gu3in1[0]);
                for (int j = 3; j < length; j++)
                {
                    if (cvwn2zy4biao3.Cells["N" + j.ToString()].Value != null && //"N"列是中古音
                        cvwn2zy4biao3.Cells["G" + j.ToString()].Value != null && //"G"列是上古音
                        cvwn2zy4biao3.Cells["N" + j.ToString()].Value.ToString().Contains(zhong1gu3vin4mu3) &&
                        vin4bu4.Any(x => cvwn2zy4biao3.Cells["G" + j.ToString()].Value.ToString().Contains(x))
                        )
                    {//第二次自改：找同上古韻部的中古同韻字
                        string vin4 = vin4bu4.First(x => cvwn2zy4biao3.Cells["G" + j.ToString()].Value.ToString().Contains(x));
                        res = ChangePronuciationOC(cvwn2zy4biao3, [vin4], shang4gu3in1, Convert.ToInt32(zhong1gu3in1[1]), ref found, exelRowsCount);
                        if (found)
                            return res;
                    }
                }

                //第三次自改： 無據試改
                res = ChangePronuciationOC(cvwn2zy4biao3, vin4bu4, shang4gu3in1, Convert.ToInt32(zhong1gu3in1[1]), ref found, exelRowsCount);
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
        
        private static Dictionary<string, Dictionary<string, int>> OnsetsOC(Worksheet ws, int exelRowsCount = 10000) //所有當前擬構的上古聲母及出現次數
        {
            Dictionary<string, Dictionary<string, int>> res = new Dictionary<string, Dictionary<string, int>>();

            for (int j = 3; j < exelRowsCount; j++)
            {
                if (ws.Cells["G" + j.ToString()].Value != null)//"G"列是上古音
                {
                    var in1zie5 = ws.Cells["G" + j.ToString()].Value.ToString();
                    string shendenhu = ws.Cells["H" + j.ToString()].Value.ToString() + ws.Cells["I" + j.ToString()].Value.ToString() + ws.Cells["J" + j.ToString()].Value.ToString();////中古聲等呼
                    string shen1 = GetOnset(in1zie5);
                    if (!res.Keys.Contains(shen1))
                        res.Add(shen1, new Dictionary<string, int>() { { shendenhu, 1 } });
                    else
                    {
                        if (!res[shen1].Keys.Contains(shendenhu))
                        {
                            res[shen1].Add(shendenhu, 1);
                        }
                        else
                        {
                            res[shen1][shendenhu]++;
                        }
                    }
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
            for (int j = 3; j < length; j++)
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
            else if (zy == "川")
            {
                zy = "𡿦";
            }
            else if (zy == "楫")
            {
                zy = "檝";
            }
            else if (zy == "斬")//須平聲同義談韻字
            {
                zy = "槧";  
            }
            else if (zy == "奭")//𥈜通赩，皆許極切。
            {
                zy = "𥈜";
            }
            return zy;
        }


        static Workbook shang4gu3duei4zhong1gu3(Worksheet ws, int length)
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
                                        if (k.EndsWith(v) || k.EndsWith(v + "h"))
                                        {
                                            int hang2hao4=0;
                                            vinmuFound=true;
                                            var vinmu = "," + k.Substring(shang4gu3sheng1mu3.Length);//e.g.: ",am", ",ʷˤəl" ...
                                            foreach (var vinbu in vin4bu4hao4)
                                            {
                                                if (vinmu.EndsWith(vinbu.Key) || vinmu.EndsWith(vinbu.Key + "h"))
                                                {
                                                    hang2hao4 = vinbu.Value;
                                                    break;
                                                }
                                            }
                                            string ss = vinmu + ws.Cells["I" + res.ToString()].Value.ToString()
                                                + ws.Cells["J" + res.ToString()].Value.ToString() + ws.Cells["K" + res.ToString()].Value.ToString();
                                            int biao3hao4 = k.Contains("ˤ") ? 1 : 0;
                                            if (wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()].Value == null ||
                                                !wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()].Value.ToString().Contains(ss))
                                            {
                                                Cell cell = wbForSave.Worksheets[biao3hao4].Cells[kv.Key.Substring(kv.Key.Length - 1) + hang2hao4.ToString()];
                                                string oldString = cell.Value?.ToString();
                                                string newString = oldString + ss;
                                                cell.PutValue(newString);
                                                if (!String.IsNullOrEmpty(oldString) && oldString.Contains(vinmu) && !oldString.Contains(vinmu + "h"))
                                                {
                                                    var characters = cell.Characters(newString.IndexOf(ss), ss.Length);
                                                    characters.Font.Color = Color.Red;
                                                }
                                            }
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
                foreach (var kv in shen1mu3duei4in4)
                {
                    wbForSave.Worksheets[0].Cells[kv.Key.Substring(kv.Key.Length - 1) + (vin4bu4hao4.Count + 1).ToString()].Value = kv.Key;
                    wbForSave.Worksheets[1].Cells[kv.Key.Substring(kv.Key.Length - 1) + (vin4bu4hao4.Count + 1).ToString()].Value = kv.Key;
                }

            }
            wbForSave.Save(@"D:\上古對中古" + DateTime.Now.ToString("yyMMddHHmm") + ".xlsx");
            return wbForSave;
        }

        /*
         * 聲母中對上。統計中古聲母來自哪些上古聲母
         */
        static Dictionary<string, Dictionary<string, int>> ShengMuZhongDueiShang(Worksheet ws, int length)
        {
            
            Dictionary<string, Dictionary<string, int>> dic = new Dictionary<string, Dictionary<string, int>>();
            int res = 3;
            try
            {
                while ((ws.Cells["G" + res.ToString()].Value == null ||
                       !String.IsNullOrWhiteSpace(ws.Cells["G" + res.ToString()].Value.ToString())) && res <= length) //"G"列是上古音
                {
                    if (ws.Cells["G" + res.ToString()].Value != null && ws.Cells["G" + res.ToString()].GetStyle().Font.Color.Name != "ffffcc00")//"G"列是上古音 ffffcc00是黃標特殊僞音須忽略
                    {
                        string in1zie5 = ws.Cells["G" + res.ToString()].Value.ToString();//"G"列是上古音
                        string shen1 = GetOnset(in1zie5);//e.g. glw
                        string den = ws.Cells["I" + res.ToString()].Value.ToString() == "三" ? "三" : "丰";
                        string shengden = ws.Cells["H" + res.ToString()].Value.ToString() + den;
                        if (dic.Keys.Contains(shengden))
                        {
                            if (dic[shengden].Keys.Contains(shen1))
                            {
                                dic[shengden][shen1]++;
                            }
                            else
                            {
                                dic[shengden].Add(shen1, 1);
                            }
                        }
                        else
                        {
                            dic.Add(shengden, new Dictionary<string, int>() { { shen1, 1 }, });
                        }
                    }
                    res++;
                }
            }
            catch (Exception ex)
            {

            }
            return dic;
        }

        private static void Huang4Üin4(List<string> denüin, string jouli, string sinli, string liuä = "")
        {
            Workbook wb = new Workbook(@"D:\《廣韻》形聲考李.xlsx");//上古表
            Worksheet ws = wb.Worksheets[0];
            var dt = ws.Cells.ExportDataTable(0, 0, 9912, 1);
            int ii;
            for (int k = 2; k < 9912; k++)
            {
                string zhongguüinmu = ws.Cells["I" + (k + 1).ToString()].Value.ToString()  + ws.Cells["J" + (k + 1).ToString()].Value.ToString() + ws.Cells["K" + (k + 1).ToString()].Value.ToString();//K列是中古韻母

                if (ws.Cells["G" + (k + 1).ToString()].Value != null)
                {
                    string shangguin = ws.Cells["G" + (k + 1).ToString()].Value.ToString();//G列是上古擬音

                    foreach (var s in denüin)
                    {
                        if (zhongguüinmu == s)
                        {
                            if (shangguin.Contains(jouli))// && !shangguin.Contains(liuä))
                            {
                                ws.Cells["G" + (k + 1).ToString()].Value = shangguin.Replace(jouli, sinli);
                            }
                        }
                    }
                }
            }
            wb.Save(@"D:\真皆真部.xlsx");

        }

        private static int Factorial(int n)
        {
            if (n == 0 || n == 1)
                return 1;
            int result = 1;
            for (int i = 2; i <= n; i++)
            {
                result *= i;
            }
            return result;
        }

        public static int CalculateCombination(int n, int k)
        {
            if (k > n || k < 0)
                throw new ArgumentException("Invalid values for n and k. k must be between 0 and n.");

            // Use the combination formula: C(n, k) = n! / (k! * (n - k)!)
            int numerator = Factorial(n);
            int denominator = Factorial(k) * Factorial(n - k);

            return numerator / denominator;
        }

        static Dictionary<string, int> ReadCsvToDictionary(string filePath)
        {
            var dict = new Dictionary<string, int>();
            using var reader = new StreamReader(filePath);
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;
                var parts = line.Split(',');
                if (parts.Length < 2) continue; // Zeile überspringen, falls weniger als 2 Spalten
                if (int.TryParse(parts[1], out int value))
                {
                    dict[parts[0]] = value;
                }
            }
            return dict;
        }

        static (int rowIndex, int rowCount) FindRowAndMergedLengthByPrefix(Worksheet ws, string prefix, int colNum)
        {
            for (int row =2; row <= ws.Cells.MaxDataRow; row++)
            {
                var cell = ws.Cells[row, colNum];
                if (cell.Value is string cellValue && cellValue.StartsWith(prefix))
                {
                    int rowCount = 1;
                    int startRow = row;
                    if (cell.IsMerged)
                    {
                        var range = cell.GetMergedRange();
                        rowCount = range.RowCount;
                        startRow = range.FirstRow;
                    }
                    // row+1, weil Aspose.Cells 0-basiert ist, Excel aber 1-basiert
                    return (startRow, rowCount);
                }
            }
            return (0, 0); // Nicht gefunden
        }

        static void CutAndInsertRows(Worksheet wsh, int startRow, int totalRows, ref int insertAtRow)
        {
            // 1. Zeilen ausschneiden (kopieren und löschen)
            //if (IsMergedAndNotFirstRow(wsh.Cells[insertAtRow,0]))
            //{

            //}
            while (IsMergedAndNotFirstRow(wsh.Cells[insertAtRow, 0]))
            {
                insertAtRow++;
            }
            wsh.Cells.InsertRows(insertAtRow, totalRows, true); // Platz schaffen
            wsh.Cells.CopyRows(wsh.Cells, insertAtRow > startRow? startRow : startRow + totalRows, insertAtRow, totalRows);

            // Wenn die Einfügeposition nach dem Ausschneidebereich liegt, muss der Startindex angepasst werden
            if (insertAtRow > startRow)
            {
                wsh.Cells.DeleteRows(startRow, totalRows, true);
            }
            else
            {
                wsh.Cells.DeleteRows(startRow + totalRows, totalRows, true);
            }
        }

        static bool IsMergedAndNotFirstRow(Cell cell)
        {
            if (cell.IsMerged)
            {
                var range = cell.GetMergedRange();
                // Prüfe, ob die aktuelle Zeile nicht die erste Zeile des Merged-Bereichs ist
                if (cell.Row > range.FirstRow)
                    return true;
            }
            return false;
        }

        static void UnmergeAndPropagateValueInColumn(Worksheet ws, int colIndex)
        {
            for (int row = 2; row <= ws.Cells.MaxDataRow; row++)
            {
                var cell = ws.Cells[row, colIndex];
                if (cell.IsMerged)
                {
                    var range = cell.GetMergedRange();
                    var originalValue = cell.Value;
                    ws.Cells.UnMerge(range.FirstRow, range.FirstColumn, range.RowCount, range.ColumnCount);
                    for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                    {
                        ws.Cells[r, colIndex].Value = originalValue;
                    }
                    row = range.FirstRow + range.RowCount - 1;
                }
            }
        }

        static void Din4Vin4Bu4(Worksheet ws, int vin4bu4lie5, int li3in1lie5)//根據上古韻母定韻部
        {
            for (int row = 2; row <= ws.Cells.MaxDataRow; row++)
            {
                var cell = ws.Cells[row, li3in1lie5];
                foreach (var vin4bu4 in liou3mang4vin4bu4)
                {
                    if (cell.Value != null && cell.Value.ToString().Contains(vin4bu4.Value))
                    {
                        ws.Cells[row, vin4bu4lie5].Value = vin4bu4.Key;
                        break;
                    }
                }                
            }
        }

        static void MoveRedCharactersToFrontInColumn(Worksheet ws, int colIndex)
        {
            for (int row = 2; row <= ws.Cells.MaxDataRow; row++)
            {
                var cell = ws.Cells[row, colIndex];
                if (cell.Value is string str && str.Length > 0)
                {
                    //var chars = cell.GetCharacters();
                    var redChars = new StringBuilder();
                    var normalChars = new StringBuilder(str);

                    // Sammle alle roten Zeichen und entferne sie aus dem Originalstring
                    for (int i = str.Length - 1; i >= 0; i--)
                    {
                        var c = cell.Characters(i, 1);
                        if (c.Font.Color.Name == "ffff0000")
                        {
                            redChars.Insert(0, str.Substring(c.StartIndex, c.Length));
                            normalChars.Remove(c.StartIndex, c.Length);
                        }
                    }

                    // Setze neuen Wert: rote Zeichen vorne, Rest dahinter
                    cell.PutValue(redChars.ToString() + normalChars.ToString());

                    // Setze die roten Zeichen wieder als rot
                    if (redChars.Length > 0)
                    {
                        var newChars = cell.Characters(0, redChars.Length);
                        newChars.Font.Color = Color.Red;
                    }
                }
            }
        }
    }
}
