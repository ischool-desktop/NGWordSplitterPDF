using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace NGWordSplitter
{
    public static class Util
    {
        /// <summary>
        /// Dear all：
        ///南港國小學生電子履歷的命名規則今日會議中已經確認，如下：
        ///2014(年)1(学期1碼)04(類别2碼)A1313xxxxx(身份号英大寫)01(案号2碼)
        ///
        ///翻譯一下。
        ///封面的檔名為：2014100身分證號01
        ///借閱紀錄檔名為：2014101身分證號01
        ///成績單檔名為：2014106身分證號01
        ///
        ///附件為身份證號、姓名、班座、學號資料。
        ///謝謝
        /// </summary>
        /// <param name="idNumber"></param>
        /// <returns></returns>
        public static string GenerateFileName(string idNumber, string typeCode)
        {
            //2014 1 00 J129999999 01
            return string.Format("20141{0}{1}01", typeCode, idNumber.ToUpper());
        }
    }

    public class IDNumberLookup
    {
        private Dictionary<string, StudentData> Lookup = new Dictionary<string, StudentData>();

        public IDNumberLookup(string fileName)
        {
            Workbook wb = new Workbook(fileName);
            Worksheet ws = wb.Worksheets[0];

            Dictionary<string, int> header_lookup = ReadColumnHeaders(ws);

            int row_offset = 1;
            for (int i = 1; i <= ws.Cells.MaxDataRow; i++)
            {
                string grade = ws.Cells[row_offset, header_lookup["年級"]].StringValue;
                string cls = ws.Cells[row_offset, header_lookup["班級"]].StringValue;
                string cn = string.Format("{0}年{1}班", grade, ConvertChinese(cls));
                string sno = ws.Cells[row_offset, header_lookup["座號"]].StringValue;
                string idnum = ws.Cells[row_offset, header_lookup["統編"]].StringValue + ""; //身分證號
                string name = ws.Cells[row_offset, header_lookup["學生"]].StringValue; //姓名
                string stunum = ws.Cells[row_offset, header_lookup["學號"]].StringValue; //姓名

                string key = GetKey(cn, sno);

                if (!Lookup.ContainsKey(key))
                    Lookup.Add(key, new StudentData() { IDNumber = idnum, Name = name, StudentNumber = stunum });

                row_offset++;
            }
        }

        private string GetKey(string cn, string sno)
        {
            return string.Format("{0}#{1}", cn, sno);
        }

        private static Dictionary<string, int> ReadColumnHeaders(Worksheet ws)
        {
            Dictionary<string, int> header_lookup = new Dictionary<string, int>();
            for (int i = 0; i <= ws.Cells.MaxDataColumn; i++)
            {
                string header = ws.Cells[0, i].StringValue;
                if (!header_lookup.ContainsKey(header))
                    header_lookup.Add(header, i);
            }
            return header_lookup;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cn">班級名稱，例：三年二班。</param>
        /// <param name="sno">座號(數字)。</param>
        /// <returns></returns>
        public StudentData GetStudentData(string cn, string sno)
        {
            string key = GetKey(cn, sno);

            if (Lookup.ContainsKey(key))
                return Lookup[key];
            else
                return null;
        }

        private string ConvertChinese(string cls)
        {
            switch (cls)
            {
                case "1":
                    return "一";
                case "2":
                    return "二";
                case "3":
                    return "三";
                case "4":
                    return "四";
                case "5":
                    return "五";
                case "6":
                    return "六";
                case "7":
                    return "七";
                case "8":
                    return "八";
                case "9":
                    return "九";
                case "10":
                    return "十";
                default:
                    return string.Empty;

            }
        }

        public class StudentData
        {
            public string Name { get; set; }

            public string IDNumber { get; set; }

            public string StudentNumber { get; set; }

            public string Info { get; set; }
        }
    }
}
