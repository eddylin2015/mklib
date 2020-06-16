using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace mklib
{
    public partial class EsStudReport
    {
        public string rate2s(Object x)
        {
            int tempi = int.Parse(x.ToString());
            if (tempi == 0 || tempi == 100) { return "   "; }
            else{ return x.ToString()+"%"; }
        }
        public string s2d(Object x,string mg="m")
        {
            if (mg == "m")
            {
                try { return String.Format("{0:0}", Decimal.Parse(x.ToString())); } catch (Exception ex) { }
                return "--";
            }
            else
            {
                try {return a2g( x); } catch (Exception ex) { }
                return "--";
            }
        }
        public string s2d2(Object x, int fixlen)
        {
            Decimal xx = Decimal.Parse(x.ToString());
            return (xx == 0) ? "   " : String.Format("{0:0.0}", xx);
        }
        public string s2d2(Object x,string mg="m")
        {
            if (mg == "m")
            {
                try
                {
                    Decimal xx = Decimal.Parse(x.ToString());
                    return (xx < 60) ? String.Format("{0:0.00}", xx) + "*" : String.Format("{0:0.00}", xx) + " ";
                }
                catch (Exception ex) { }
                return "--";
            }
            else
            {
                try
                {
                    Decimal xx = Decimal.Parse(x.ToString());

                    return a2g(xx);
                }
                catch (Exception ex) { }
                return "--";
            }
        }
        public static string crs2s(string x, int c)
        {
            /*
            foreach (EncodingInfo ei in Encoding.GetEncodings())
            {
                Encoding e = ei.GetEncoding();

                Console.Write("{0,-15}", ei.CodePage);
                if (ei.CodePage == e.CodePage)
                    Console.Write("    ");
                else
                    Console.Write("*** ");

                Console.Write("{0,-25}", ei.Name);
                if (ei.CodePage == e.CodePage)
                    Console.Write("    ");
                else
                    Console.Write("*** ");

                Console.Write("{0,-25}", ei.DisplayName);
                if (ei.CodePage == e.CodePage)
                    Console.Write("    ");
                else
                    Console.Write("*** ");
                Console.Write("{0}_{1}_{2}",x,x.Length,e.GetByteCount(x));
            }*/
            System.Text.Encoding encodeUTF8 = System.Text.Encoding.UTF8;
            int utf7_cnt = encodeUTF8.GetByteCount(x);
            int tempint = x.Length *2- ( x.Length * 3- utf7_cnt )/2;            
            for (int i = tempint; i < c; i++) x += " ";
            return x;
        }
        public static string px2s(Object o)
        {
            string x = o.ToString();
            if (x.ToString().Equals("1")) return "[*]";
            if (x.ToString().Equals("2")) return "[P]";
            return "   ";
        }
        public static string class2n(string txt)
        {
            txt = txt.Replace("P", "小");
            txt = txt.Replace("SC", "高");
            txt = txt.Replace("SG", "初");
            txt = txt.Replace("1", "一");
            txt = txt.Replace("2", "二");
            txt = txt.Replace("3", "三");
            txt = txt.Replace("4", "四");
            txt = txt.Replace("5", "五");
            txt = txt.Replace("6", "六");
            txt = txt.Replace("A", "信");
            txt = txt.Replace("B", "望");
            txt = txt.Replace("C", "愛");
            txt = txt.Replace("D", "善");
            txt = txt.Replace("E", "樂");
            return txt;
        }

        public static string _a2g(Decimal m, String pclass)
        {
            if (pclass.ToUpper().StartsWith('P'))
            {
                if (m >= 95) { return "A "; }
                else if (m >= 90) { return "A-"; }
                else if (m >= 85) { return "B+"; }
                else if (m >= 80) { return "B "; }
                else if (m >= 75) { return "B-"; }
                else if (m >= 70) { return "C+"; }
                else if (m >= 65) { return "C "; }
                else if (m >= 60) { return "C-"; }
                else { return "D "; }
            }
            else
            {
                if (m == 100) { return "A+"; }
                else if (m >= 95) { return "A "; }
                else if (m >= 90) { return "A-"; }
                else if (m >= 85) { return "B+"; }
                else if (m >= 80) { return "B "; }
                else if (m >= 75) { return "B-"; }
                else if (m >= 70) { return "C+"; }
                else if (m >= 65) { return "C "; }
                else if (m >= 61) { return "C-"; }
                else if (m == 60) { return "D "; }
                else { return "F "; }
            }
        }
        public static string hr = "-------------------------------------------------------------------------------------------------------------\n";

        public static string PGCTxt(DataRow cRow)
        {
            return string.Format("{0}{1,5}{2,5}{3,8}{4,5}{5,5}{6,8}{7,5}{8,5}{9,8}{10,16}{11,14}\n",
                crs2s(cRow["GC_Name"].ToString(), 22),
                "", "", cRow["grade1"],
                "", "", cRow["grade2"],
                "", "", cRow["grade3"],
                "",
                ""
                );
        }
        public static String SCDTxt(DataRow x)
        {
            return "";
        }
        public  String PCDTxt(DataRow cRow)
        {
            return string.Format("{0}{1,5}{2,5}{3,8}{4,5}{5,5}{6,8}{7,5}{8,5}{9,8}{10,16}{11,14}\n",
                            crs2s("             總平均成績", 22),
                            "", "", s2d2(cRow["mark1"]),
                            "", "", s2d2(cRow["mark2"]),
                            "", "", s2d2(cRow["mark3"]),
                            s2d2(cRow["mark"]),
                            ""
                            );
        }

        public static string GetCDFMT(DataRow dr, String cno, String term)
        {
            if(cno.StartsWith("P")){
                return String.Format("遲到: {0,3}次  缺席:  {1,3}節  曠課: {2,3}節\n操行:   {4}  違紀:  {5,3}印  褒獎: {6,3}印",
                dr["wrg_later" + term],
                dr["wrg_absence" + term],
                dr["wrg_truancy_t" + term],
                dr["wrg_truancy_s" + term].ToString(),
                crs2s(dr["conduct" + term].ToString(), 3),
                dr["WrgMarks" + term],
                int.Parse(dr["honor" + term].ToString()) +
                int.Parse(dr["wrg_later" + term].ToString()));
            }
            else{
            return String.Format("遲到: {0,3}次  缺席:  {1,3}節  曠課: {2,3}節{3,3}次\n操行:   {4}  違紀:  {5,3}印  褒獎: {6,3}印",
                dr["wrg_later" + term],
                dr["wrg_absence" + term],
                dr["wrg_truancy_t" + term],
                dr["wrg_truancy_s" + term].ToString(),
                crs2s(dr["conduct" + term].ToString(), 3),
                dr["WrgMarks" + term],
                int.Parse(dr["honor" + term].ToString()) +
                int.Parse(dr["wrg_later" + term].ToString()));
            }
        }

    }
}
