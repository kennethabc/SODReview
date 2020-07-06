using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace SODReview
{
   
    public partial class Form1 : Form
    {
        
        string[] teamList = new string[] {"HR",
"GL_oversea",
"GL_PC",
"IT EIS",
"TR oversea",
"AR",
"CAO",
"GL_PM",
"Procurement",
"AP&MAP",
"RPA",
"IT EAS",
"GL_PS",
"LEGAL",
"CF",
"AP_TW",
"TR China",
"CEO",
"NON_SSCGZ",
"MDM",
"TAX",
"CIO",
"Agency",
"Treasury"};
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            

          

            
        }
        private void GetData(string mappingPath, string ouput)
        {
            if (!Directory.Exists("VEGA"))
            {
                
                Directory.CreateDirectory("VEGA");
            }
            DataRal sqlHelpr = new DataRal("Server=10.180.3.40\\HKHKGIDBSSC0018;Database=EAS_DB;User Id=EAS_Kingdee_admin;Pwd=E@s2015Kingdee");
            DataTable DT_VegaUser = sqlHelpr.ExecutePS("YS_P_R_UserList '','','','SSC'");
            ExcelHelper.ExcelHelper exhelper = new ExcelHelper.ExcelHelper(mappingPath);
            Console.WriteLine(mappingPath);
            DataTable DT_CNUser = exhelper.ExcelToDataTable("Sheet1", true);

            var CNUser = DT_CNUser.AsEnumerable();
            var VegaUser = DT_VegaUser.AsEnumerable();

            var query = from vega in VegaUser
                        join cn in CNUser
                        on vega.Field<string>("Short Name").ToLower() equals cn.Field<string>("User Code")
                        into temp
                        from cn in temp.DefaultIfEmpty(DT_CNUser.NewRow())
                        select new
                        {
                            userid = vega.Field<string>("Short Name"),
                            userName = vega.Field<string>("员工名字"),
                            email = vega.Field<string>("femail"),
                            userAccessCompany = vega.Field<string>("组织名字"),
                            userGroup = vega.Field<string>("Group"),
                            userCompany = vega.Field<string>("用户组"),
                            Department = cn.Field<string>("Department") == null ? "na" : cn.Field<string>("Department")
                        };

            DataTable res = new DataTable();
            res.Columns.Add("LionID");
            res.Columns.Add("UserName");
            res.Columns.Add("FEmail");
            res.Columns.Add("UserAccessCompany");
            res.Columns.Add("UserGroup");
            res.Columns.Add("UserCompany");
            res.Columns.Add("Department");



            foreach (var item in query)
            {
                res.Rows.Add(new string[] { item.userid, item.userName, item.email, item.userAccessCompany, item.userGroup, item.userCompany, item.Department });


            }
            ExcelHelper.ExcelHelper output = null;
            foreach (string item in teamList)
            {
                DataTable dt = GetSubDataTable(res, "Department='"+item+"'");
                if (dt != null)
                {
                    output = new ExcelHelper.ExcelHelper("VEGA\\"+"" + item + "_VEGA.xlsx");
                    Console.WriteLine(item);
                    output.DataTableToExcel(dt, "Sheet1", true);
                }

            }
            MessageBox.Show("已完成VEGA");

          
        }

        private void GetRFData(string mappingPath, string outputFolder)
        {
            if (!Directory.Exists("Readsoft"))
            {
                Directory.CreateDirectory("Readsoft");
            }
            DataRal sqlHelpr = new DataRal("Server=10.180.4.221;Database=KAP_DATA;User Id=SP-adm;Pwd=Fe8mCPjf");
            DataTable DT_VegaUser = sqlHelpr.ExecutePS("sp_user_list '','',''");
            ExcelHelper.ExcelHelper exhelper = new ExcelHelper.ExcelHelper(mappingPath);
            Console.WriteLine(mappingPath);
            DataTable DT_CNUser = exhelper.ExcelToDataTable("Sheet1", true);

            var CNUser = DT_CNUser.AsEnumerable();
            var VegaUser = DT_VegaUser.AsEnumerable();

            var query = from vega in VegaUser
                        join cn in CNUser
                        on vega.Field<string>("User_Account").Replace("ll\\", "").ToLower() equals cn.Field<string>("User Code")
                        into temp
                        from cn in temp.DefaultIfEmpty(DT_CNUser.NewRow())
                        select new
                        {
                            userid = vega.Field<string>("User_Account"),
                            userName = vega.Field<string>("UserName"),
                            email = vega.Field<string>("email_addr"),
                            userAccessCompany = vega.Field<string>("Division_Name"),
                            NextApprover = vega.Field<string>("NextApprover"),
                            AccessGroup = vega.Field<string>("Name"),
                            ApprovalProcess = vega.Field<string>("Approval_Process"),
                            ApprovalLimit = vega.Field<double>("Approval_Limit").ToString(),
                            Department = cn.Field<string>("Department")
                        };

            DataTable res = new DataTable();
            res.Columns.Add("User_Account");
            res.Columns.Add("UserName");
            res.Columns.Add("FEmail");
            res.Columns.Add("UserAccessCompany");
            res.Columns.Add("NextApprover");
            res.Columns.Add("AccessGroup");
            res.Columns.Add("ApprovalProcess");
            res.Columns.Add("Approval_Limit");
            res.Columns.Add("Department");

            foreach (var item in query)
            {
                res.Rows.Add(new string[] {
                    item.userid,
                    item.userName,
                    item.email,
                    item.userAccessCompany,
                    item.NextApprover,
                    item.AccessGroup,
                    item.ApprovalProcess,
                    item.ApprovalLimit,
                    item.Department
                }); 


            }
            ExcelHelper.ExcelHelper output = null;
            foreach (string item in teamList)
            {
                DataTable dt = GetSubDataTable(res, "Department='" + item + "'");
                if (dt != null)
                {
                    Console.WriteLine(item);
                    output = new ExcelHelper.ExcelHelper("Readsoft\\"+"" + item + "_Readsoft.xlsx");
                    output.DataTableToExcel(dt, "Sheet1", true);
                }
                

            }
            MessageBox.Show("已完成Readsoft");

        }

        private void GetAPCData(string mappingPath, string APCUserList)
        {
            if (!Directory.Exists("APC"))
            {
                Directory.CreateDirectory("APC");
            }
            ExcelHelper.ExcelHelper mapEx = new ExcelHelper.ExcelHelper(mappingPath);
            ExcelHelper.ExcelHelper apcEX = new ExcelHelper.ExcelHelper(APCUserList);
            DataTable mapDt = mapEx.ExcelToDataTable("Sheet1", true);
            DataTable apcDt = apcEX.ExcelToDataTable("AP Central User List", true);

            var query = from apc in apcDt.AsEnumerable()
                        join map in mapDt.AsEnumerable() on apc.Field<string>("User ID").ToLower().Replace("ll\\","") equals map.Field<string>("User Code")
                        into temp
                        from t in temp.DefaultIfEmpty(mapDt.NewRow())
                        select new
                        {
                            Country = apc.Field<string>("Country"),
                            CompanyName = apc.Field<string>("Company Name"),
                            UserID = apc.Field<string>("User ID"),
                            UserName = apc.Field<string>("User Name"),
                            Email = apc.Field<string>("Email"),
                            FunctionGroup = apc.Field<string>("Function Group"),
                            Group = apc.Field<string>("Group"),
                            Department = t.Field<string>("Department")
                        };

            DataTable res = new DataTable();
            res.Columns.Add("Country");
            res.Columns.Add("CompanyName");
            res.Columns.Add("UserID");
            res.Columns.Add("UserName");
            res.Columns.Add("Email");
            res.Columns.Add("FunctionGroup");
            res.Columns.Add("Group");
            res.Columns.Add("Department");

            foreach (var item in query)
            {
                res.Rows.Add(new string[] { item.Country, item.CompanyName, item.UserID, 
                    item.UserName, item.Email, item.FunctionGroup, item.Group, item.Department });
            }

            ExcelHelper.ExcelHelper output = null;
            foreach (string item in teamList)
            {
                DataTable dt = GetSubDataTable(res, "Department='" + item + "' and Group='Re:Sources' and Country='Hong Kong'");
                if (dt != null)
                {
                    output = new ExcelHelper.ExcelHelper("APC\\"+"" + item + "_AP Central.xlsx");
                    Console.WriteLine(item);
                    output.DataTableToExcel(dt, "Sheet1", true);
                }

            }
            MessageBox.Show("已完成APC");
        }

        private void GetSPData(string mappingPath, string SPUserList)
        {
            if (!Directory.Exists("Spectra"))
            {
                Directory.CreateDirectory("Spectra");
            }
            ExcelHelper.ExcelHelper mapEx = new ExcelHelper.ExcelHelper(mappingPath);
            ExcelHelper.ExcelHelper apcEX = new ExcelHelper.ExcelHelper(SPUserList);
            DataTable mapDt = mapEx.ExcelToDataTable("Sheet1", true);
            DataTable apcDt = apcEX.ExcelToDataTable("Spectra Active User Listing", true);

            var query = from apc in apcDt.AsEnumerable()
                        join map in mapDt.AsEnumerable() on apc.Field<string>("ShortName").ToLower() equals map.Field<string>("User Code")
                        into temp
                        from t in temp.DefaultIfEmpty(mapDt.NewRow())
                        select new
                        {
                            Country = apc.Field<string>("Country"),
                            CompanyName = apc.Field<string>("BusinessUnitIDName1"),
                            DataBase = apc.Field<string>("Database"),
                            UserID = apc.Field<string>("ShortName"),
                            UserName = apc.Field<string>("Name1"),
                            Email = apc.Field<string>("Email"),
                            UserGroup = apc.Field<string>("UserGroupIdName1"),
                            BelongTo = apc.Field<string>("Belong to"),
                            Department = t.Field<string>("Department")

                        };

            DataTable res = new DataTable();
            res.Columns.Add("Country");
            res.Columns.Add("CompanyName");
            res.Columns.Add("DataBase");
            res.Columns.Add("UserID");
            res.Columns.Add("UserName");
            res.Columns.Add("Email");
            res.Columns.Add("UserGroup");
            res.Columns.Add("BelongTo");
            res.Columns.Add("Department");

            foreach (var item in query)
            {
                res.Rows.Add(new string[] { item.Country, item.CompanyName,item.DataBase, item.UserID,
                    item.UserName, item.Email, item.UserGroup, item.BelongTo, item.Department }) ;
            }

            ExcelHelper.ExcelHelper output = null;
            foreach (string item in teamList)
            {
                DataTable dt = GetSubDataTable(res, "Department='" + item + "' and BelongTo='Re:sources' and Country='Hong Kong'");
                if (dt != null)
                {
                    output = new ExcelHelper.ExcelHelper("Spectra\\"+"" + item + "_Spectra.xlsx");
                    Console.WriteLine(item);
                    output.DataTableToExcel(dt, "Sheet1", true);
                }

            }
            MessageBox.Show("已完成Spectra");
        }

        private DataTable GetSubDataTable(DataTable MainDt, string keyword)
        {

            DataRow[] rows = MainDt.Select(keyword);
            Console.WriteLine(rows.Length);
            if (rows.Length > 0)
            {
                DataTable dt = rows[0].Table.Clone();

                foreach (DataRow row in rows)
                {
                    DataRow newRow = dt.NewRow();
                    newRow.ItemArray = row.ItemArray;
                    dt.Rows.Add(newRow);
                }
                return dt;

            }

            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel | *.xlsx";
            ofd.Title = "Select Mapping File";
            ofd.InitialDirectory = @"C:\Users\"+Environment.UserName+@"\Desktop";
            ofd.ShowDialog();
            

            if (ofd.CheckFileExists)
            {
                this.textBox1.Text = ofd.FileName;
            }
            else
            {
                MessageBox.Show("Target File is not existed", "File Not Exists");
            }

        }



        private void button2_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                
                Thread t = new Thread(()=>GetData(this.textBox1.Text, ""));
                t.Start();

                
                //GetData(this.textBox1.Text, this.textBox2.Text);
            }
            if (checkBox2.Checked)
            {
                Thread t = new Thread(() => GetRFData(this.textBox1.Text, ""));
                //GetRFData(this.textBox1.Text, this.textBox2.Text);
                t.Start();
            }
            if (checkBox3.Checked)
            {
                Thread t = new Thread(() => GetAPCData(this.textBox1.Text, this.aptxt.Text));
                t.Start();
                //GetAPCData(this.textBox1.Text, this.aptxt.Text);
            }
            if (checkBox4.Checked)
            {
                Thread t = new Thread(() => GetSPData(this.textBox1.Text, this.sptxt.Text));
                t.Start();
                //GetSPData(this.textBox1.Text, this.sptxt.Text);
            }
        }


        private void checkBox3_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                panel2.Visible = true;
            }
            else
            {
                panel2.Visible = false;
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel | *.xlsx";
            ofd.Title = "Select UserList File";
            ofd.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop";
            ofd.ShowDialog();
            aptxt.Text = ofd.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel | *.xlsx";
            ofd.Title = "Select UserList File";
            ofd.InitialDirectory = @"C:\Users\" + Environment.UserName + @"\Desktop";
            ofd.ShowDialog();
            sptxt.Text = ofd.FileName;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                panel1.Visible = true;
            }
            else
            {
                panel1.Visible = false;
            }
        }
    }
}
