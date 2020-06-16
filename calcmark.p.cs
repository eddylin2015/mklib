using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace mklib
{   public interface iEsStudReport
    {
        string endofcult(DataRow cds, String term, String mg);
        string endofprof(DataRow cds, String term, String mg);
        string endofcourse(DataRow cds, String term, String mg);
        string fmt_course_term(DataRow mark, String term, String mg);
        string fmt_subcourse_term(DataRow mark, String term, String mg);
        string itera(String term, String mg, DataRow[] marks, DataRow[] cds);
        string MrkTxt(String term, String mg, DataRow[] drs, DataRow[] cds, DataRow[] acs, DataRow[] gcs);
        string ACTFMT(DataRow[] cRow);
        string GCFMT(DataRow[] cRow);
        string a2g(Object m);
        string a2g(Decimal m);
    }
    public partial class EsStudReport:iEsStudReport
    {       
        public virtual string endofcult(DataRow cds,String term,String mg) { return null; }
        public virtual string endofprof(DataRow cds, String term, String mg) { return null; }
        public virtual string endofcourse(DataRow cds, String term, String mg) { return null; }
        public virtual string fmt_course_term(DataRow mark,String term,String mg) { return null; }
        public virtual string fmt_subcourse_term(DataRow mark, String term, String mg) { return null; }
        public virtual string itera(String term,String mg,DataRow[] marks,DataRow[] cds)
        {
            var res = "";
            String ctype = null;
            var techcourse = false;
            for (var j = 0; j < marks.Length; j++)
            {
                if (ctype == null)
                {
                    ctype = marks[j]["c_t_type"].ToString();
                    res += cttypedisc(ctype);
                    if (ctype == "職業文化"|| ctype == "職業專業") techcourse = true; 
                }
                else if (!ctype.Equals(marks[j]["c_t_type"].ToString()))
                {
                    if (ctype == "職業文化")
                    {
                        res += endofcult(cds[0], term, mg);
                        techcourse = true;
                    }
                    ctype = marks[j]["c_t_type"].ToString();
                    res += cttypedisc(ctype);
                }
                if (j == (marks.Length - 1)
                    || marks[j]["groupid"].ToString().Equals("100")
                    || !marks[j]["groupid"].ToString().Equals(marks[j + 1]["groupid"].ToString()))
                {
                    res += fmt_course_term(marks[j], term, mg);
                }
                else
                {
                    res += fmt_subcourse_term(marks[j], term, mg);
                }
            }
            if (techcourse)
            {
                res += endofprof(cds[0], term, mg);
            }
            else
            {
                res += endofcourse(cds[0], term, mg);
            }
            return res;
        }
        protected virtual string cttypedisc(String cttype)
        {
            var t = "";
            if (cttype == "必修") { t = "基礎學科"; }
            else if (cttype == "必選") { t = "拓展及自選學科"; }
            else if (cttype == "職業文化") { t = "社會文化學科"; }
            else if (cttype == "職業專業") { t = "專業科技及實踐學科"; }
            else { t = cttype; }
            return "[" + t + "]\n";
        }
        public virtual String MrkTxt(String term, String mg, DataRow[] drs, DataRow[] cds, DataRow[] acs , DataRow[] gcs)
        {
            return null;
        }
        public virtual String ACTFMT(DataRow[] cRow)
        {
            return null;
        }
        public virtual String GCFMT(DataRow[] cRow)
        {
            return null;
        }
        public virtual string a2g(Object m) { return "--"; }
        public virtual string a2g(Decimal m) { return "--"; }
    }
    public class SEsStudReportFillCDS : EsStudReport
    {
        public override string a2g(Object m)
        {
            Decimal m0 = Decimal.Parse(m.ToString());
            return EsStudReport._a2g(m0,"S");
        }
        public override string a2g(Decimal m)
        {
            return EsStudReport._a2g(m,"S");
        }
        private string hr = "-----------------------------------------------------------------------------------------------------------------------\n";
        /*
         * 12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
         * -----------------------------------------------------------------------------------------------------------------------
         * 1-------------------221--4-1--4-1-----7--1--4-1--4-1-----7--1--4-1--4-1-----7------1-----7---------1-3-------
         */
        private string fmtstr1 = "{0}{1,4} {2,4} {3,7}    {4,4} {5,4} {6,7}    {7,4} {8,4} {9,7}      {10,7}        {11,3}      {12,4}\n";
        private string fmtstr2 =             "{0}{1,7}    {2,4} {3,4} {4,7}    {5,4} {6,4} {7,7}      {8,3}        {9,3}      {10,4}\n";

        public override String MrkTxt(String term, String mg, DataRow[] marks, DataRow[] cds, DataRow[] acs, DataRow[] gcs)
        {
            string marktxt = itera(term, mg, marks, cds);
            marktxt += ACTFMT(acs);
            marktxt += tlwrg(cds, acs);
            return marktxt;
        }
        
        public string tlwrg(DataRow[] cds,DataRow[] acs)
        {
            decimal x = 0;
            decimal y = 0;
            decimal z = 0;            
            try { x = decimal.Parse(cds[0]["total_crs_ncp"].ToString()); } catch (Exception e) { }
            if(acs!=null&&acs.Length>0) try { y = decimal.Parse(acs[0]["addXF"].ToString()); } catch (Exception e) { }
            try { z = decimal.Parse(cds[0]["volunteer_hr"].ToString()); } catch (Exception e) { }
            string fmt1 = "                                                                                     學生卓越表現附加學分{0,13}\n";
            string fmt2 = "                                                                                   全學年合計扣減學分總數{0,13}\n";
            string fmt3 = "全年義工服務時數： {0} 小時。\n";
            string res = String.Format(fmt1, y.ToString("0.0")) + String.Format(fmt2, (y + x).ToString("0.0")) +(z==0?"": String.Format(fmt3, z));
            return res;
        }
        public override String ACTFMT(DataRow[] cRow)
        {
            if(cRow.Length>0){
            return string.Format(fmtstr1,
                crs2s("課外活動" ,25), 
                "", "", a2g(cRow[0]["grade1"]) + " ",
                "", "", a2g(cRow[0]["grade2"]) + " ",
                "", "", a2g(cRow[0]["grade3"]) + " ",
                "",
                "",""
                );}else{return "\n";}
        }
        public override string fmt_course_term(DataRow mark, String term, String mg)
        {
            //Console.WriteLine(mark["coursename"].ToString());
            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            if(rate < 100 && mark["coursename"].ToString().IndexOf("(")<0) frate =" "+( rate < 100 ? (rate + "%") : "");
            return string.Format(fmtstr1,
                        crs2s("  "+mark["coursename"].ToString() + frate, 25),
                        s2d(mark["t1"]), s2d(mark["e1"]), s2d2(mark["total1"]),
                        s2d(mark["t2"]), s2d(mark["e2"]), s2d2(mark["total2"]),
                        s2d(mark["t3"]), s2d(mark["e3"]), s2d2(mark["total3"]),
                        s2d2(mark["total"]),
                        px2s(mark["P_X"]),
                        s2d2(mark["sub_c_p"], 1)
                        );
        }
        public override string fmt_subcourse_term(DataRow mark, String term, String mg)
        {

            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            if(rate < 100 && mark["coursename"].ToString().IndexOf("(")<0) frate =" "+( rate < 100 ? (rate + "%") : "");
            return string.Format(fmtstr1,
                        crs2s("  " + mark["coursename"].ToString() + frate, 25),
                        s2d(mark["t1"]), s2d(mark["e1"]), "",
                        s2d(mark["t2"]), s2d(mark["e2"]), "",
                        s2d(mark["t3"]), s2d(mark["e3"]), "",
                        "",
                        px2s(mark["P_X"]),
                        s2d2(mark["sub_c_p"],1));
        }
        public override string endofcourse(DataRow cRow, String term, String mg) {
            
            return hr+string.Format(fmtstr1,
                          crs2s("         總平均成績", 25),
                          "", "", s2d2(cRow["mark1"]),
                          "", "", s2d2(cRow["mark2"]),
                          "", "", s2d2(cRow["mark3"]),
                          s2d2(cRow["mark"]),
                          "",""
                          );
        }
        public override String GCFMT(DataRow[] cRow)
        {
            return string.Format(fmtstr1,
               crs2s(cRow[0]["GC_Name"].ToString(), 25),
               "", "", cRow[0]["grade1"] +" ",
               "", "", cRow[0]["grade2"] + " ",
               "", "", cRow[0]["grade3"] + " ",
               "",
               "",""
               );
        }


    }
    public class PEsStudReportFillCDS : EsStudReport
    {
        
        public override string a2g(Object m)
        {
            Decimal m0 = Decimal.Parse(m.ToString());
            return EsStudReport._a2g(m0, "P");
        }
        public override string a2g(Decimal m)
        {
            return EsStudReport._a2g(m, "P");
        }

        private string hr = "------------------------------------------------------------------------------------------------------------\n";
        /*
         * 1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
         * -------------------------------------------------------------------------------------------------------------
         * 1-------------------221--4-1--4-1-----7--1--4-1--4-1-----7--1--4-1--4-1-----7------1-----7---------1-3-------
         */
        private string fmtstr1 = "{0}{1,4} {2,4} {3,7}  {4,4} {5,4} {6,7}  {7,4} {8,4} {9,7}      {10,7}           {11,3}\n";
        private string fmtstr2 =             "{0}{1,7}  {2,4} {3,4} {4,7}  {5,4} {6,4} {7,7}      {8,3}\n";
        private string fmtactstr = "{0}{1,4}  {2,4} {3,4} {4,7}  {5,4} {6,4} {7,7}      {8,7}           {9,3}\n";
        protected override string cttypedisc(String cttype)
        {
            return "";
        }
        public override String MrkTxt(String term, String mg, DataRow[] marks, DataRow[] cds, DataRow[] acs, DataRow[] gcs)
        {
            
            String marktxt = itera(term, mg, marks, cds);
            marktxt += ACTFMT(acs);
            marktxt += GCFMT(gcs);
            return marktxt;
        }
        public override string endofcult(DataRow cds, String term, String mg) { return null; }
        public override string endofprof(DataRow cds, String term, String mg) { return null; }
        public override string fmt_course_term(DataRow mark, String term, String mg) {
            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            return string.Format(fmtstr1,
                        crs2s(mark["coursename"].ToString() + frate, 22),
                        s2d(mark["t1"]), s2d(mark["e1"]), s2d2(mark["total1"]),
                        s2d(mark["t2"]), s2d(mark["e2"]), s2d2(mark["total2"]),
                        s2d(mark["t3"]), s2d(mark["e3"]), s2d2(mark["total3"]),
                        s2d2(mark["total"]),
                        px2s(mark["P_X"]));
        }
        public override string fmt_subcourse_term(DataRow mark, String term, String mg)
        {
            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            return string.Format(fmtstr1,
                        crs2s(mark["coursename"].ToString() + frate, 22),
                        s2d(mark["t1"]), s2d(mark["e1"]), "",
                        s2d(mark["t2"]), s2d(mark["e2"]), "",
                        s2d(mark["t3"]), s2d(mark["e3"]), "",
                        "",
                        px2s(mark["P_X"]));
        }
        private static String Act2S(string x, int c)
        {
            System.Text.Encoding encodeUTF8 = System.Text.Encoding.UTF8;
            int utf7_cnt = encodeUTF8.GetByteCount(x);
            int tempint = x.Length * 2 - (x.Length * 3 - utf7_cnt) / 2;
            String pattern = "[ a-zA-Z]+";
            System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex(pattern);
            int _c = rgx.Matches(x).Count;
            c = tempint > 57 ? (c + 57*2-3 ) : c;
            for (int i = tempint; i < (c - _c); i++) x += " ";
            return x;
        }
        public override String ACTFMT(DataRow[] cRow)
        {
            if(cRow.Length>0){
                //string x = string.Format("活動課組別-{0}\n", cRow[0]["act_py"]);
                string x = string.Format("活動課-{0}", cRow[0]["act_py"]);
                return string.Format(fmtactstr,
               Act2S(x, 35),
               a2g(cRow[0]["grade1"])+" ","", "", a2g(cRow[0]["grade2"]) + " ", "", "", a2g(cRow[0]["grade3"]) + " ",
               "", ""
               );
            }else{
                return "";
            }
        }
        public override String GCFMT(DataRow[] cRow)
        {
            if (cRow.Length == 0) return "";
            return
            string.Format(fmtstr2,
                          crs2s(cRow[0]["GC_Name"].ToString(), 32),
                          crs2s(cRow[0]["grade1"].ToString(),3),
                          "", "", crs2s(cRow[0]["grade2"].ToString(), 3),
                          "", "", crs2s(cRow[0]["grade3"].ToString(), 3),
                          "",
                          ""
                          );
        }
        public override string endofcourse(DataRow cRow, String term, String mg)
        {
            return hr+ string.Format(fmtstr1,
                      crs2s("            總平均成績", 22),
                      "", "", s2d2(cRow["mark1"]),
                      "", "", s2d2(cRow["mark2"]),
                      "", "", s2d2(cRow["mark3"]),
                      s2d2(cRow["mark"]),
                      "");
        }
    }
    public class VEsStudReportFillCDS : SEsStudReportFillCDS
    {
        private string hr = "-----------------------------------------------------------------------------------------------------------------------\n";
        private string fmtstr1 = "{0}{1,4} {2,4} {3,7}   {4,4} {5,4} {6,7}   {7,4} {8,4} {9,7}       {10,6}          {11,3}\n";
        private string fmtstr2 = "{0}{1,4} {2,4} {3,7}   {4,4} {5,4} {6,7}   {7,4} {8,4} {9,7}       {10,6}          {11,3}\n";
        /* 
  數學                           C-   F      F       C-   F      D       C-   F      F             F           [P]         
  公民教育                       B+   A-     A-      B    B+     B+      B    B      B             B+                     
  體育                           C+   C      C+      C+   C+     C+      D    D      D             C                      
  歷史                           B    B+     B       B+   B      B+      B-   B      B-            B                      
  地理                           B    B      B       C-   B-     C+      C-   A-     B-            B-                     
                學科總平均成績               C+                  C+                  C             C+           C+ 
-----------------------------------------------------------------------------------------------------------------------
123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890        
         */       
        public override string fmt_course_term(DataRow mark, String term, String mg)
        {
            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            return string.Format(fmtstr1,
                        crs2s("  " + mark["coursename"].ToString() + frate, 31),
                        s2d(mark["t1"], mg), s2d(mark["e1"], mg), s2d2(mark["total1"],mg),
                        s2d(mark["t2"], mg), s2d(mark["e2"], mg), s2d2(mark["total2"], mg),
                        s2d(mark["t3"], mg), s2d(mark["e3"], mg), s2d2(mark["total3"], mg),
                        s2d2(mark["total"], mg),
                        px2s(mark["P_X"]));
        }
        public override string fmt_subcourse_term(DataRow mark, String term, String mg)
        {
            int rate = int.Parse(mark["rate"].ToString());
            string frate = rate < 100 ? (rate + "%") : "";
            return string.Format(fmtstr1,
                        crs2s("  " + mark["coursename"].ToString() + frate, 31),
                        s2d(mark["t1"], mg), s2d(mark["e1"], mg), "",
                        s2d(mark["t2"], mg), s2d(mark["e2"], mg), "",
                        s2d(mark["t3"], mg), s2d(mark["e3"], mg), "",
                        "",
                        px2s(mark["P_X"]));
        }
        private string muefmt(object x,object y,string mg="m")
        {
            try
            {
                Decimal x1 = Decimal.Parse(x.ToString());
                Decimal x2 = Decimal.Parse(y.ToString());
                if (x2 > x1) {
                    if (mg.Equals("m")) { return string.Format("{0:0.00}", x2); }
                    else if (mg.Equals("g")) { return a2g(x2); }
                }
            }
            catch(Exception e)
            {
            }
            return "";
        }
        public override string endofcult(DataRow cRow,string term,string mg) {
            return string.Format(fmtstr1,
                         crs2s("            總平均成績", 31),
                         "", "", s2d2(cRow["voca_cult_avg1"], mg),
                         "", "", s2d2(cRow["voca_cult_avg2"], mg),
                         "", "", s2d2(cRow["voca_cult_avg3"], mg),
                         s2d2(cRow["voca_cult_avg"], mg),
                         muefmt(cRow["voca_cult_avg"],cRow["voca_cult_mue"], mg)
                         ) +hr;
        }
        public override string endofprof(DataRow cRow, string term, string mg) {
            return string.Format(fmtstr1,
                          crs2s("            總平均成績", 31),
                          "", "", s2d2(cRow["voca_prof_avg1"], mg),
                          "", "", s2d2(cRow["voca_prof_avg2"], mg),
                          "", "", s2d2(cRow["voca_prof_avg3"], mg),
                          s2d2(cRow["voca_prof_avg"], mg),
                          muefmt(cRow["voca_prof_avg"],cRow["voca_prof_mue"], mg)
                          ) +hr;
        }
        public override String ACTFMT(DataRow[] cRow)
        {
            return string.Format(fmtstr1,
                crs2s("課外活動", 31),
                "", "", a2g(cRow[0]["grade1"]) + " ",
                "", "", a2g(cRow[0]["grade2"]) + " ",
                "", "", a2g(cRow[0]["grade3"]) + " ",
                "",
                "", ""
                );
        }
        public override string MrkTxt(String term, String mg, DataRow[] marks, DataRow[] cds, DataRow[] acs, DataRow[] gcs)
        {            
            string marktxt = itera(term, mg, marks, cds);
            marktxt += ACTFMT(acs);
            return marktxt;
        }
    }

 }
