using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Text.RegularExpressions;
using System.IO;
using NPOI.XWPF.UserModel;

namespace mklib
{
    public delegate void delegate_output_JSON(DataSet ds, string pclassno, string term, string mg, string sess, string session, string pdate, string filter,TextWriter sw, bool namemask = false);

    struct MK_Stud_Total_REC
    {
        public int gcnt;
        public decimal m1;
        public decimal m2;
        public decimal m3;
        public decimal m;
        public decimal total_crs_ncp;
        public int voca_cult_gcnt;
        public decimal voca_cult_avg1;
        public decimal voca_cult_avg2;
        public decimal voca_cult_avg3;
        public decimal voca_cult_avg;
        public decimal voca_cult_mue;
        public int voca_prof_gcnt;
        public decimal voca_prof_avg1;
        public decimal voca_prof_avg2;
        public decimal voca_prof_avg3;
        public decimal voca_prof_avg;
        public decimal voca_prof_mue;
        public int allpass1;
        public int allpass2;
        public int allpass3;
    }
    public class MK_Utils
    {
        public static String str2html(String txt)
        {
            return txt.Replace("\n", "<br>").Replace("\"", "&quot;").Replace("\r", "").Replace(" ", "&nbsp;");
        }
        public static int NoOfConduct(String conduct)
        {
            string cd = conduct.Trim().Replace("＋", "+").Replace("－", "-").Replace("一", "-");
            String rule = "丙-,丙,丙+,乙-,乙,乙+,甲-,甲,甲+";
            string[] ra = rule.Split(',');
            for (int i = 0; i < ra.Length; i++)
            {
                if (ra[i].Equals(cd)) return i;
            }
            return 0;
        }
        public static String GetEval(String cno, int allpass, decimal mark, string conduct, int later, int absence, int truancy, out int EvalAddHonorInt)
        {
            cno = cno.ToUpper();
            string[] HonorEvalDESC_ARR = { "品學兼優生\n", "品行優異生\n", "學業優異生\n", "勤學生\n" };
            int[] HonorEvalAdjMark_ARR = { 3, 2, 2, 1 };
            int[] SMg = { 75, 70 };
            int[] PMg = { 85, 75 };
            int mg1 = 85;
            int mg2 = 75;
            if (cno[0] == 'P')
            {
                mg1 = PMg[0]; mg2 = PMg[1];
                HonorEvalAdjMark_ARR[0] = 0;
                HonorEvalAdjMark_ARR[1] = 0;
                HonorEvalAdjMark_ARR[2] = 0;
                HonorEvalAdjMark_ARR[3] = 0;
            }
            else if (cno[0] == 'S') { mg1 = SMg[0]; mg2 = SMg[1]; }
            String res = "";
            EvalAddHonorInt = 0;
            int noOfcond = MK_Utils.NoOfConduct(conduct);
            if (allpass == 1)
            {
                if (mark >= mg1 && noOfcond >= 6)
                {
                    res += HonorEvalDESC_ARR[0];
                    EvalAddHonorInt = HonorEvalAdjMark_ARR[0];
                }
                else if (noOfcond >= 6)
                {
                    res += HonorEvalDESC_ARR[1];
                    EvalAddHonorInt = HonorEvalAdjMark_ARR[1];
                }
                else if (mark >= mg1 && noOfcond >= 4)
                {
                    res += HonorEvalDESC_ARR[2];
                    EvalAddHonorInt = HonorEvalAdjMark_ARR[2];
                }
                else if (mark >= mg2 && noOfcond >= 4)
                {
                    res += HonorEvalDESC_ARR[3];
                    EvalAddHonorInt = HonorEvalAdjMark_ARR[3];
                }
            }
            if (later == 0 && absence == 0 && truancy == 0)
            {
                res += "全勤生";
            }
            return res;
        }
    }
    public partial class CalcMark
    {
        private static string _appPath = null;
        public static string AppPath
        {
            get
            {
                if (_appPath == null)
                    _appPath = System.IO.Path.GetDirectoryName(
                        System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Substring(6);
                return _appPath;
            }
        }
        public static string year;

        private static void CalcGroup(List<DataRow> dr, ref MK_Stud_Total_REC st, Hashtable ngtbl)
        {
            decimal gm1 = 0M;
            decimal gm2 = 0M;
            decimal gm3 = 0M;
            decimal gm = 0M;
            decimal gm_mue = 0M;
            DataRow lastrow = null;
            foreach (DataRow r in dr)
            {
                lastrow = r;
                r["total1"] = 0;
                r["total2"] = 0;
                r["total3"] = 0;
                r["total"] = 0;
                r["VOCA_MUE"] = 0;
                r["sub_c_p"] = 0;
                r["eog"] = 0;
                decimal tmpm1 = (decimal.Parse(r["t1"].ToString()) + decimal.Parse(r["e1"].ToString())) / 2;
                decimal tmpm2 = (decimal.Parse(r["t2"].ToString()) + decimal.Parse(r["e2"].ToString())) / 2;
                decimal tmpm3 = (decimal.Parse(r["t3"].ToString()) + decimal.Parse(r["e3"].ToString())) / 2;
                decimal tmpm = tmpm1 * 0.3M + tmpm2 * 0.3M + tmpm3 * 0.4M;
                if (tmpm < 60 && decimal.Parse(r["pk"].ToString()) >= 60) { r["P_X"] = 2; }
                else if (tmpm < 60) { r["P_X"] = 1; } else { r["P_X"] = 0; }
                gm1 += tmpm1 * int.Parse(r["rate"].ToString()) / 100;
                gm2 += tmpm2 * int.Parse(r["rate"].ToString()) / 100;
                gm3 += tmpm3 * int.Parse(r["rate"].ToString()) / 100;
                gm += tmpm * int.Parse(r["rate"].ToString()) / 100;
                if (tmpm < 60 && tmpm < decimal.Parse(r["pk"].ToString()))
                {
                    gm_mue += decimal.Parse(r["pk"].ToString());
                }
                else
                {
                    gm_mue += tmpm * int.Parse(r["rate"].ToString()) / 100;
                }
            }
            gm1 = Math.Round(gm1, 2, MidpointRounding.AwayFromZero);
            gm2 = Math.Round(gm2, 2, MidpointRounding.AwayFromZero);
            gm3 = Math.Round(gm3, 2, MidpointRounding.AwayFromZero);
            gm = Math.Round(gm, 2, MidpointRounding.AwayFromZero);
            gm_mue = Math.Round(gm_mue, 2, MidpointRounding.AwayFromZero);
            lastrow["total1"] = gm1;
            lastrow["total2"] = gm2;
            lastrow["total3"] = gm3;
            lastrow["total"] = gm;
            lastrow["VOCA_MUE"] = gm_mue;
            lastrow["eog"] = 1;
            if (gm >= 60)
            {
                foreach (DataRow r in dr)
                {
                    r["P_X"] = 0;
                }
            }
            else if (gm < 60)
            {
                foreach (DataRow r in dr)
                {
                    if (int.Parse(r["P_X"].ToString()) == 1)
                    {
                        lastrow["sub_c_p"] = ngtbl[lastrow["c_ng_id"].ToString()];
                    }
                }
            }
            st.m1 += gm1;
            st.m2 += gm2;
            st.m3 += gm3;
            st.m += gm;
            st.gcnt += 1;
            if (gm1 >= 60) st.allpass1++;
            if (gm2 >= 60) st.allpass2++;
            if (gm >= 60) st.allpass3++;
            if (gm < 60)
            {
                st.total_crs_ncp += decimal.Parse(lastrow["sub_c_p"].ToString());
            }
            if (lastrow["c_t_type"].ToString().Equals("職業文化"))
            {
                st.voca_cult_avg1 += gm1;
                st.voca_cult_avg2 += gm2;
                st.voca_cult_avg3 += gm3;
                st.voca_cult_avg += gm;
                st.voca_cult_mue += gm_mue;
                st.voca_cult_gcnt += 1;
            }
            if (lastrow["c_t_type"].ToString().Equals("職業專業"))
            {
                st.voca_prof_avg1 += gm1;
                st.voca_prof_avg2 += gm2;
                st.voca_prof_avg3 += gm3;
                st.voca_prof_avg += gm;
                st.voca_prof_mue += gm_mue;
                st.voca_prof_gcnt += 1;
            }
        }
        private static void Calc(DataSet ds, Hashtable ngtbl, string cno)
        {
            DataTable tb = ds.Tables["Table"];
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                MK_Stud_Total_REC st = new MK_Stud_Total_REC();
                int groupid = 0;
                int curr_groupid = 0;
                List<DataRow> group_courses = new List<DataRow>();
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_mk"))
                {                    
                    if (groupid == 0)
                    {
                        curr_groupid = int.Parse(cRow["groupid"].ToString());
                        groupid = curr_groupid;
                        group_courses.Add(cRow);
                    }
                    else
                    {
                        curr_groupid = int.Parse(cRow["groupid"].ToString());
                        if (groupid == 100)
                        {
                            CalcGroup(group_courses, ref st, ngtbl);
                            group_courses.Clear();
                            group_courses.Add(cRow);
                            groupid = curr_groupid;
                        }
                        else if (groupid != curr_groupid)
                        {
                            CalcGroup(group_courses, ref st, ngtbl);
                            group_courses.Clear();
                            group_courses.Add(cRow);
                            groupid = curr_groupid;
                        }
                        else if (groupid == curr_groupid)
                        {
                            group_courses.Add(cRow);
                        }
                    }
                }
                CalcGroup(group_courses, ref st, ngtbl);
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_cd"))
                {
                    cRow["mark1"] = Math.Round(st.m1 / st.gcnt, 2, MidpointRounding.AwayFromZero);
                    cRow["mark2"] = Math.Round(st.m2 / st.gcnt, 2, MidpointRounding.AwayFromZero);
                    cRow["mark3"] = Math.Round(st.m3 / st.gcnt, 2, MidpointRounding.AwayFromZero);
                    cRow["mark"] = Math.Round(st.m / st.gcnt, 2, MidpointRounding.AwayFromZero);
                    cRow["total_crs_ncp"] = st.total_crs_ncp;
                    cRow["allpass1"] = st.gcnt - st.allpass1 == 0 ? 1 : 0;
                    cRow["allpass2"] = st.gcnt - st.allpass2 == 0 ? 1 : 0;
                    cRow["allpass3"] = st.gcnt - st.allpass3 == 0 ? 1 : 0;
                    int EvalHonor = 0;
                    cRow["SchoolEval1"] = MK_Utils.GetEval(
                        cno,
                        st.gcnt - st.allpass1 == 0 ? 1 : 0,
                        decimal.Parse(cRow["mark1"].ToString()),
                        cRow["conduct1"].ToString(),
                        int.Parse(cRow["wrg_later1"].ToString()),
                        int.Parse(cRow["wrg_absence1"].ToString()),
                        int.Parse(cRow["wrg_truancy_t1"].ToString()),
                        out EvalHonor);
                    cRow["SE_HONOR1"] = EvalHonor;
                    cRow["SchoolEval2"] = MK_Utils.GetEval(
                        cno,
                        st.gcnt - st.allpass2 == 0 ? 1 : 0,
                        decimal.Parse(cRow["mark2"].ToString()),
                        cRow["conduct2"].ToString(),
                        int.Parse(cRow["wrg_later2"].ToString()),
                        int.Parse(cRow["wrg_absence2"].ToString()),
                        int.Parse(cRow["wrg_truancy_t2"].ToString()),
                        out EvalHonor);
                    cRow["SE_HONOR2"] = EvalHonor;
                    cRow["SchoolEval3"] = MK_Utils.GetEval(
                        cno,
                        st.gcnt - st.allpass3 == 0 ? 1 : 0,
                        decimal.Parse(cRow["mark"].ToString()),
                        cRow["conduct3"].ToString(),
                        int.Parse(cRow["wrg_later3"].ToString()),
                        int.Parse(cRow["wrg_absence3"].ToString()),
                        int.Parse(cRow["wrg_truancy_t3"].ToString()),
                        out EvalHonor);
                    cRow["SE_HONOR3"] = EvalHonor;
                    if (st.voca_cult_gcnt > 0)
                    {
                        cRow["voca_cult_avg1"] = Math.Round(st.voca_cult_avg1 / st.voca_cult_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_cult_avg2"] = Math.Round(st.voca_cult_avg2 / st.voca_cult_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_cult_avg3"] = Math.Round(st.voca_cult_avg3 / st.voca_cult_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_cult_avg"] = Math.Round(st.voca_cult_avg / st.voca_cult_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_cult_mue"] = Math.Round(st.voca_cult_mue / st.voca_cult_gcnt, 2, MidpointRounding.AwayFromZero);
                    }
                    if(st.voca_prof_gcnt>0)
                    { 
                        cRow["voca_prof_avg1"] = Math.Round(st.voca_prof_avg1 / st.voca_prof_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_prof_avg2"] = Math.Round(st.voca_prof_avg2 / st.voca_prof_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_prof_avg3"] = Math.Round(st.voca_prof_avg3 / st.voca_prof_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_prof_avg"] = Math.Round(st.voca_prof_avg / st.voca_prof_gcnt, 2, MidpointRounding.AwayFromZero);
                        cRow["voca_prof_mue"] = Math.Round(st.voca_prof_mue / st.voca_prof_gcnt, 2, MidpointRounding.AwayFromZero);
                    }
                }
            }
        }
        static List<string> genQR(DataSet ds)
        {
            List<string> res = new List<string>();
            DataTable tb = ds.Tables["Table"];
            String t_flag = "total";
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                String total = "0000";
                String vca = "0000";
                String vpa = "0000";
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_cd"))
                {
                    if (decimal.Parse(cRow["mark3"].ToString()) > 0)
                    {
                        total = cRow["mark"].ToString().Replace(".", "");
                        vca = cRow["voca_cult_mue"].ToString().Replace(".", "");
                        vpa = cRow["voca_prof_mue"].ToString().Replace(".", "");
                        t_flag = "total";
                    }
                    else if (decimal.Parse(cRow["mark2"].ToString()) > 0)
                    {
                        total = cRow["mark2"].ToString().Replace(".", "");
                        vca = cRow["voca_cult_avg2"].ToString().Replace(".", "");
                        vpa = cRow["voca_prof_avg2"].ToString().Replace(".", "");
                        t_flag = "total2";
                    }
                    else if (decimal.Parse(cRow["mark1"].ToString()) > 0)
                    {
                        total = cRow["mark1"].ToString().Replace(".", "");
                        vca = cRow["voca_cult_avg1"].ToString().Replace(".", "");
                        vpa = cRow["voca_prof_avg1"].ToString().Replace(".", "");
                        t_flag = "total1";
                    }
                }
                List<int> gmstr = new List<int>();
                int gmpk = 0;
                int groupid = 0;
                int curr_groupid = 0;
                List<DataRow> group_courses = new List<DataRow>();
                int row_cnt = 0;
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_mk"))
                {
                    row_cnt++;
                    if (cRow["P_X"].ToString() == "1")
                    {
                        gmpk += 1 << (row_cnt - 1);
                    }
                    if (groupid == 0)
                    {
                        curr_groupid = int.Parse(cRow["groupid"].ToString());
                        groupid = curr_groupid;
                        group_courses.Add(cRow);
                    }
                    else
                    {
                        curr_groupid = int.Parse(cRow["groupid"].ToString());
                        if (groupid == 100)
                        {
                            //CalcGroup
                            {
                                int lastgm = int.Parse(group_courses[group_courses.Count - 1][t_flag].ToString()[0].ToString());
                                gmstr.Add(lastgm);
                            }
                            group_courses.Clear();
                            group_courses.Add(cRow);
                            groupid = curr_groupid;
                        }
                        else if (groupid != curr_groupid)
                        {
                            //CalcGroup
                            {
                                int lastgm = int.Parse(group_courses[group_courses.Count - 1][t_flag].ToString()[0].ToString());
                                gmstr.Add(lastgm);
                            }
                            group_courses.Clear();
                            group_courses.Add(cRow);
                            groupid = curr_groupid;
                        }
                        else if (groupid == curr_groupid)
                        {
                            group_courses.Add(cRow);
                        }
                    }
                }
                //CalcGroup
                {
                    int lastgm = int.Parse(group_courses[group_courses.Count - 1][t_flag].ToString()[0].ToString());
                    gmstr.Add(lastgm);
                }
                String _qrml = "";
                int _qrmi = 0;
                foreach (int ele in gmstr)
                {
                    _qrml += ele.ToString();
                    _qrmi += ele;
                }
                _qrmi = _qrmi % gmstr.Count;
                String seat = tb.Rows[i]["curr_seat"].ToString();
                if (seat.Length == 1) seat = "0" + seat;
                string classno = tb.Rows[i]["curr_class"].ToString();
                if (classno.Length == 3) classno = classno + "R";
                String filename = String.Format("{0}{1}{2}{3}", year, tb.Rows[i]["curr_class"], seat, tb.Rows[i]["dsej_ref"]);
                if (classno[classno.Length - 1] == 'E') { total = vca + vpa; }
                String qrstr = String.Format("{0}{1}{2}{3}{4}{5}{6}{7:X}", year, classno, seat, tb.Rows[i]["dsej_ref"], total, _qrml, Convert.ToChar(65 + _qrmi), gmpk);
                res.Add(String.Format("{0};{1}", filename, qrstr));
            }
            return res;
        }


        public static void outputTXT(DataSet ds, string pclassno, string term, string mg, string sess, string session, string pdate, string filter, TextWriter sw,bool namemask = false)
        {
            string pattern = @"\d+";
            Regex rx = new Regex(pattern);
            MatchCollection matches = rx.Matches(filter);
            EsStudReport esrp = null;
            if (pclassno.StartsWith("P")) { esrp = new PEsStudReportFillCDS(); }
            else if (pclassno.EndsWith("E"))
            {
                esrp = new VEsStudReportFillCDS();
            }
            else { esrp = new SEsStudReportFillCDS(); }

            DataTable tb = ds.Tables["Table"];
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                string seat = tb.Rows[i]["curr_seat"].ToString();
                if (matches.Count > 0)
                {
                    int crt = 0;
                    for (crt = 0; crt < matches.Count; crt++)
                    {
                        if (matches[crt].Value.Equals(seat)) break;
                    }
                    if (crt >= matches.Count) continue;
                }
                String marktxt = esrp.MrkTxt(term, mg, tb.Rows[i].GetChildRows("sr_mk"),
                    tb.Rows[i].GetChildRows("sr_cd"),
                    tb.Rows[i].GetChildRows("sr_ac"),
                    tb.Rows[i].GetChildRows("sr_gc")
                    );
                Console.WriteLine(marktxt);
            }
        }

        public static void outputDOCX(DataSet ds, string pclassno, string term, string mg, string sess, string session, string pdate, string filter, TextWriter sw,bool namemask = false)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("session", typeof(String));
            dt.Columns.Add("p_date", typeof(String));
            dt.Columns.Add("dsej_ref", typeof(String));
            dt.Columns.Add("C_NAME", typeof(String));
            dt.Columns.Add("E_NAME", typeof(String));
            dt.Columns.Add("classname", typeof(String));
            dt.Columns.Add("curr_class", typeof(String));
            dt.Columns.Add("seat", typeof(String));
            dt.Columns.Add("stud_ref", typeof(String));
            dt.Columns.Add("Markdetail", typeof(String));
            dt.Columns.Add("conduct", typeof(String));
            dt.Columns.Add("schooleval1", typeof(String));
            dt.Columns.Add("pingyu1", typeof(String));
            dt.Columns.Add("conduct2", typeof(String));
            dt.Columns.Add("schooleval2", typeof(String));
            dt.Columns.Add("pingyu2", typeof(String));
            dt.Columns.Add("conduct3", typeof(String));
            dt.Columns.Add("schooleval3", typeof(String));
            dt.Columns.Add("pingyu3", typeof(String));
            dt.Columns.Add("act_py", typeof(String));
            dt.Columns.Add("qrbc", typeof(String));
            dt.Columns.Add("photo", typeof(String));
            EsStudReport esrp = null;
            if (pclassno.StartsWith("P")) { esrp = new PEsStudReportFillCDS(); }
            else if (pclassno.EndsWith("E")) { esrp = new VEsStudReportFillCDS(); }
            else { esrp = new SEsStudReportFillCDS(); }
            String classno = "";
            DataTable tb = ds.Tables["Table"];
            for (int i = 0; i <  tb.Rows.Count; i++)
            {
                DataRow drr = dt.NewRow();  //選擇輸出:1,2,3,4
                string seat = tb.Rows[i]["curr_seat"].ToString();
                classno = tb.Rows[i]["curr_class"].ToString();
                drr["dsej_ref"] = tb.Rows[i]["dsej_ref"];
                drr["C_NAME"] = tb.Rows[i]["c_name"];
                drr["E_NAME"] = tb.Rows[i]["e_name"];
                drr["curr_class"] = tb.Rows[i]["curr_class"];
                drr["classname"] = EsStudReport.class2n(classno);
                drr["seat"] = tb.Rows[i]["curr_seat"];
                drr["stud_ref"] = tb.Rows[i]["stud_ref"];
                drr["session"] = session;
                drr["p_date"] = pdate;
                String marktxt = esrp.MrkTxt(term, mg, tb.Rows[i].GetChildRows("sr_mk"),
                   tb.Rows[i].GetChildRows("sr_cd"),
                   tb.Rows[i].GetChildRows("sr_ac"),
                   tb.Rows[i].GetChildRows("sr_gc")
                   );
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_cd"))
                {
                    drr["conduct"] = EsStudReport.GetCDFMT(cRow, pclassno, "1");
                    drr["conduct2"] = EsStudReport.GetCDFMT(cRow, pclassno, "2");
                    drr["conduct3"] = EsStudReport.GetCDFMT(cRow, pclassno, "3");
                    drr["schooleval1"] = cRow["SchoolEval1"];
                    drr["schooleval2"] = cRow["SchoolEval2"];
                    drr["schooleval3"] = cRow["SchoolEval3"];
                }
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_ac"))
                {
                    drr["act_py"] = cRow["act_py"];
                }
                drr["Markdetail"] = marktxt;
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_py"))
                {
                    drr["pingyu1"] = cRow["py1"];
                    drr["pingyu2"] = cRow["py2"];
                    drr["pingyu3"] = cRow["py3"];
                }
                drr["photo"] = @"c:\photo" + sess + @"\" + drr["dsej_ref"] + ".jpg";
                drr["qrbc"] = @"c:\photo" + sess + @"\qrcode\" + sess + drr["curr_class"] + (seat.Length == 1 ? "0" : "") + seat + drr["dsej_ref"] + ".jpg";
                dt.Rows.Add(drr);
            }
            //////////////
            string templ_page1 = AppPath + @"\xmldata" + @"\T1516.xml";
            if (classno.StartsWith("P")) templ_page1 = AppPath + @"\xmldata" + @"\PT1920.xml";// + @"\PT1415.xml";
            if (classno.EndsWith("E"))
            {
                templ_page1 = mg.Equals("m") ? AppPath + @"\xmldata" + @"\VOCAMT1516.xml" : AppPath + @"\xmldata" + @"\VOCAGT1516.xml";
            }
            string out_filename = AppPath + @"\xmlout" + @"\MarkTable_" + classno + "_" + mg + ".doc";
            FileInfo finfo = new FileInfo(out_filename);
            if (finfo.Exists  &&   (DateTime.Now - finfo.CreationTime).TotalMinutes < 2)
            {
                Console.WriteLine("Exists Docx{0} {1}", out_filename,finfo.CreationTime);
            } else {
                Console.WriteLine("Make Docx{0}", out_filename);
                mklib.TDB2WML dw = new mklib.TDB2WML(new mklib.TParseWMLitem(null), dt, templ_page1);
                dw.makeDoc(out_filename);
            }
            sw.Write(out_filename);
        }
        static void Cell_NOPI_memo(String memo_str, String ReplaceTxt, XWPFTableCell cell)
        {
            string[] li = memo_str.Split("\n");
            int max_line = cell.Paragraphs.Count;
            for (int mi = 0; mi < max_line; mi++)
            {
                if (cell.Paragraphs[mi].Text.Equals(ReplaceTxt))
                    cell.Paragraphs[mi].ReplaceText(ReplaceTxt, mi < li.Length ? li[mi] : "");
            }
        }

        public static void outputDOCXX(DataSet ds, string pclassno, string term, string mg, string sess, string session, string pdate, string filter, TextWriter sw, bool namemask = false)
        {
            string temple_file = AppPath + @"\xmldata\NPOI_T1516.docx";
            if (pclassno.ToUpper().StartsWith('P')) temple_file = AppPath + @"\xmldata\NPOI_PT1920.docx";
            if (pclassno.EndsWith("E"))
            {
                temple_file = mg.Equals("m") ? AppPath + @"\xmldata" + @"\NPOI_VOCAMT1516.docx" : AppPath + @"\xmldata" + @"\NPOI_VOCAGT1516.docx";
            }
            string out_filename = AppPath + @"\xmlout" + @"\MarkTable_" + pclassno + "_" + mg + ".docx";

            Stream fs = File.OpenRead(temple_file);
            XWPFDocument doc = new XWPFDocument(fs);
            
            mklib.EsStudReport esrp = null;
            if (pclassno.StartsWith("P")) { esrp = new mklib.PEsStudReportFillCDS(); }
            else if (pclassno.EndsWith("E"))
            {
                esrp = new mklib.VEsStudReportFillCDS();
            }
            else { esrp = new mklib.SEsStudReportFillCDS(); }
            if (filter == null) filter = "";
            string pattern = @"\d+";
            Regex rx = new Regex(pattern);
            MatchCollection matches = rx.Matches(filter);
            DataTable tb = ds.Tables["Table"];
            int page_no = 0;
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                string seat = tb.Rows[i]["curr_seat"].ToString();
                string classno = tb.Rows[i]["curr_class"].ToString();
                string dsej_ref = tb.Rows[i]["dsej_ref"].ToString();
                string curr_class = tb.Rows[i]["curr_class"].ToString();
                string classname = mklib.EsStudReport.class2n(classno);
                string stud_ref = tb.Rows[i]["stud_ref"].ToString();
                if (matches.Count > 0)
                {
                    int crt = 0;
                    for (crt = 0; crt < matches.Count; crt++)
                    {
                        if (matches[crt].Value.Equals(seat)) break;
                    }
                    if (crt >= matches.Count) continue;
                }
                String marktxt = esrp.MrkTxt(term, mg, tb.Rows[i].GetChildRows("sr_mk"),
                    tb.Rows[i].GetChildRows("sr_cd"),
                    tb.Rows[i].GetChildRows("sr_ac"),
                    tb.Rows[i].GetChildRows("sr_gc")
                );
                int j = page_no * 4; page_no++;
                if (j >= doc.Tables.Count) { Console.WriteLine("ERROR doc tables are not enought!"); break; }
                var _tbl = doc.Tables[j];

                string photo_path = @"c:\photo" + sess + @"\" + dsej_ref + ".jpg";
                string qrbc_path = @"c:\photo" + sess + @"\qrcode\" + sess + curr_class + (seat.Length == 1 ? "0" : "") + seat + dsej_ref + ".jpg";
                if (!File.Exists(photo_path)) photo_path = @"c:/photo1920/none.jpg";
                if (!File.Exists(qrbc_path)) qrbc_path = @"c:/photo1920/qrcode/none.jpg";
                FileStream gfs = null;
                {
                    _tbl.Rows[0].GetCell(1).Paragraphs[0].ReplaceText("#PLphoto|", "");
                    gfs = new FileStream(photo_path, FileMode.Open, FileAccess.Read);
                    XWPFRun gr = _tbl.Rows[0].GetCell(1).Paragraphs[0].CreateRun();
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1120000, 1380000);
                    gfs.Close();
                }
                {
                    doc.Tables[j + 3].Rows[0].GetCell(3).Paragraphs[0].ReplaceText("#PLqrbc|", "");
                    gfs = new FileStream(qrbc_path, FileMode.Open, FileAccess.Read);
                    XWPFRun gr = doc.Tables[j + 3].Rows[0].GetCell(3).Paragraphs[0].CreateRun();
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1480000, 1480000);
                    gfs.Close();
                }
                _tbl.Rows[1].GetCell(1).Paragraphs[0].ReplaceText("#TLsession|", session);
                _tbl.Rows[1].GetCell(3).Paragraphs[0].ReplaceText("#TLp_date|", pdate);
                _tbl.Rows[1].GetCell(5).Paragraphs[0].ReplaceText("#TLdsej_ref|", tb.Rows[i]["dsej_ref"].ToString());
                _tbl.Rows[2].GetCell(1).Paragraphs[0].ReplaceText("#TLC_NAME|", tb.Rows[i]["C_NAME"].ToString());
                _tbl.Rows[2].GetCell(1).Paragraphs[0].ReplaceText("#TLE_NAME|", tb.Rows[i]["E_NAME"].ToString());
                _tbl.Rows[3].GetCell(1).Paragraphs[0].ReplaceText("#TLclassname|", mklib.EsStudReport.class2n(pclassno));
                _tbl.Rows[3].GetCell(3).Paragraphs[0].ReplaceText("#TLseat|", seat);
                _tbl.Rows[3].GetCell(5).Paragraphs[0].ReplaceText("#TLstud_ref|", tb.Rows[i]["stud_ref"].ToString());
                //#LSMarkdetai
                var _tblm = doc.Tables[j + 1];

                Cell_NOPI_memo(marktxt, "#LSMarkdetail|", _tblm.Rows[2].GetCell(0));

                var _tblp = doc.Tables[j + 2];
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_cd"))
                {
                    string conduct = mklib.EsStudReport.GetCDFMT(cRow, pclassno, "1");
                    string conduct2 = mklib.EsStudReport.GetCDFMT(cRow, pclassno, "2");
                    string conduct3 = mklib.EsStudReport.GetCDFMT(cRow, pclassno, "3");
                    string schooleval1 = cRow["SchoolEval1"].ToString();
                    string schooleval2 = cRow["SchoolEval2"].ToString();
                    string schooleval3 = cRow["SchoolEval3"].ToString();

                    Cell_NOPI_memo(conduct, "#LQconduct|", _tblp.Rows[1].GetCell(1));
                    Cell_NOPI_memo(conduct2, "#LQconduct2|", _tblp.Rows[2].GetCell(1));
                    Cell_NOPI_memo(conduct3, "#LQconduct3|", _tblp.Rows[3].GetCell(1));
                    Cell_NOPI_memo(schooleval1, "#LQschooleval1|", _tblp.Rows[1].GetCell(2));
                    Cell_NOPI_memo(schooleval2, "#LQschooleval2|", _tblp.Rows[2].GetCell(2));
                    Cell_NOPI_memo(schooleval3, "#LQschooleval3|", _tblp.Rows[3].GetCell(2));
                }
                if (pclassno.ToUpper().StartsWith('S'))
                    foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_ac"))
                    {
                        _tblp.Rows[4].GetCell(1).Paragraphs[0].ReplaceText("#TLact_py|", cRow["act_py"].ToString());
                    }
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_py"))
                {
                    _tblp.Rows[1].GetCell(3).Paragraphs[0].ReplaceText("#TLpingyu1|", cRow["py1"].ToString());
                    _tblp.Rows[2].GetCell(3).Paragraphs[0].ReplaceText("#TLpingyu2|", cRow["py2"].ToString());
                    _tblp.Rows[3].GetCell(3).Paragraphs[0].ReplaceText("#TLpingyu3|", cRow["py3"].ToString());
                }
            }
            for (int i = doc.Tables.Count - 1; i >= page_no * 4; i--)
            {
                var _t = doc.Tables[i];
                doc.RemoveBodyElement(doc.GetPosOfTable(_t));
            }
            using (FileStream _sw = File.Create(out_filename))
            {
                doc.Write(_sw);
            }
            sw.Write(out_filename);
        }

        public static void outputJSON(DataSet ds, string pclassno, string term, string mg, string sess, string session, string pdate, string filter, TextWriter sw,bool namemask = false)
        {
            string pattern = @"\d+";
            Regex rx = new Regex(pattern);
            MatchCollection matches = rx.Matches(filter);
            sw.WriteLine("[");
            DataTable tb = ds.Tables["Table"];
            int rowcnt = 0;
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                //選擇輸出:1,2,3,4
                string seat = tb.Rows[i]["curr_seat"].ToString();
                if (matches.Count > 0)
                {
                    int crt = 0;
                    for (crt = 0; crt < matches.Count; crt++)
                    {
                        if (matches[crt].Value.Equals(seat)) break;
                    }
                    if (crt >= matches.Count) continue;
                }

                if (rowcnt > 0) { sw.WriteLine(","); }
                rowcnt++;
                sw.WriteLine("{");
                for (int j = 0; j < tb.Columns.Count; j++)
                {
                    if (j > 0) { sw.WriteLine(","); }
                    if (namemask && (tb.Columns[j].ColumnName.Equals("stud_ref") || tb.Columns[j].ColumnName.Equals("dsej_ref") || tb.Columns[j].ColumnName.Equals("c_name") || tb.Columns[j].ColumnName.Equals("e_name") || tb.Columns[j].ColumnName.Equals("curr_class")))
                    {
                        sw.Write("{0}:\"{1}\"", tb.Columns[j].ColumnName, "---");
                    }
                    else
                    {
                        sw.Write("{0}:\"{1}\"", tb.Columns[j].ColumnName, tb.Rows[i][j]);
                    }
                }
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_py"))
                {
                    for (int j = 0; j < cRow.Table.Columns.Count; j++)
                    {
                        if (!cRow.Table.Columns[j].ColumnName.Equals("stud_ref"))
                            sw.Write(",{0}:\"{1}\"", cRow.Table.Columns[j].ColumnName, MK_Utils.str2html(cRow[j].ToString()));
                    }
                }
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_cd"))
                {
                    for (int j = 0; j < cRow.Table.Columns.Count; j++)
                    {
                        if (!cRow.Table.Columns[j].ColumnName.Equals("stud_ref"))
                            if (cRow.Table.Columns[j].DataType.Name.Equals("Decimal")
                                || cRow.Table.Columns[j].DataType.Name.Equals("Int32")
                                || cRow.Table.Columns[j].DataType.Name.Equals("Int16")
                                || cRow.Table.Columns[j].DataType.Name.Equals("SByte")
                                )
                            {
                                sw.Write(",{0}:{1}", cRow.Table.Columns[j].ColumnName, cRow[j]);
                            }
                            else
                            {
                                sw.Write(",{0}:\"{1}\"", cRow.Table.Columns[j].ColumnName, MK_Utils.str2html(cRow[j].ToString()));
                            }
                    }
                }
                sw.WriteLine(",\"marks\":[");
                int rid = 0;
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_mk"))
                {
                    if (rid > 0) { sw.WriteLine(","); }
                    sw.WriteLine("{");
                    for (int j = 0; j < cRow.Table.Columns.Count; j++)
                    {
                        if (j > 0) { sw.WriteLine(","); }
                        if (cRow.Table.Columns[j].DataType.Name.Equals("Decimal"))
                        {
                            sw.Write("{0}:{1}", cRow.Table.Columns[j].ColumnName, cRow[j]);
                        }
                        else
                        {
                            sw.Write("{0}:\"{1}\"", cRow.Table.Columns[j].ColumnName, cRow[j]);
                        }
                    }
                    sw.WriteLine("}");
                    rid++;
                }
                sw.WriteLine("]");
                sw.WriteLine(",\"acmarks\":[");
                rid = 0;
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_ac"))
                {
                    if (rid > 0) { sw.WriteLine(","); }
                    sw.WriteLine("{");
                    for (int j = 0; j < cRow.Table.Columns.Count; j++)
                    {
                        if (j > 0) { sw.WriteLine(","); }
                        sw.Write("{0}:\"{1}\"", cRow.Table.Columns[j].ColumnName, MK_Utils.str2html(cRow[j].ToString()));
                    }
                    sw.WriteLine("}");
                    rid++;
                }
                sw.WriteLine("]");
                sw.WriteLine(",\"gcmarks\":[");
                rid = 0;
                foreach (DataRow cRow in tb.Rows[i].GetChildRows("sr_gc"))
                {
                    if (rid > 0) { sw.WriteLine(","); }
                    sw.WriteLine("{");
                    for (int j = 0; j < cRow.Table.Columns.Count; j++)
                    {
                        if (j > 0) { sw.WriteLine(","); }
                        sw.Write("{0}:\"{1}\"", cRow.Table.Columns[j].ColumnName, MK_Utils.str2html(cRow[j].ToString()));
                    }
                    sw.WriteLine("}");
                    rid++;
                }
                sw.WriteLine("]");

                sw.WriteLine("}");
            }
            sw.WriteLine("]");
        }
    }
}

