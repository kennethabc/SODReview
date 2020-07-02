using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace SODReview
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            

          

            
        }
        private void GetData(string mappingPath, string ouput)
        {
            DataRal sqlHelpr = new DataRal("Server=10.180.3.40\\HKHKGIDBSSC0018;Database=EAS_DB;User Id=EAS_Kingdee_admin;Pwd=E@s2015Kingdee");
            DataTable DT_VegaUser = sqlHelpr.ExecutePS("YS_P_R_UserList '','','','SSC'");
            ExcelHelper.ExcelHelper exhelper = new ExcelHelper.ExcelHelper(mappingPath);
            Console.WriteLine(mappingPath);
            DataTable DT_CNUser = exhelper.ExcelToDataTable("Sheet1", true);

            var CNUser = DT_CNUser.AsEnumerable();
            var VegaUser = DT_VegaUser.AsEnumerable();

            var query = from vega in VegaUser
                        join cn in CNUser
                        on vega.Field<string>("Short Name") equals cn.Field<string>("User Code")
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
            res.Columns.Add("Manager");



            foreach (var item in query)
            {
                res.Rows.Add(new string[] { item.userid, item.userName, item.email, item.userAccessCompany, item.userGroup, item.userCompany, item.Department });


            }
            string savePath = "";
            if (ouput != "")
            {
                savePath = ouput + @"\VEGA.xlsx";
            }
            else
            {
                savePath = "VEGA.xlsx";
            }
            ExcelHelper.ExcelHelper ex = new ExcelHelper.ExcelHelper(savePath);
            ex.DataTableToExcel(res, "Sheet1", true);
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

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = true;
            fbd.ShowDialog();
            fbd.Description = "Please select the save path";
            textBox2.Text = fbd.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                
                GetData(textBox1.Text, textBox2.Text);
            }
            else
            {
                MessageBox.Show("Please select mapping first", "Mapping select first");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process  p = new Process();
            p.StartInfo.FileName = "explorer";
            
            if (this.textBox2.Text != null)
            {
                p.StartInfo.Arguments = @"/select," + this.textBox2.Text;
                
            }
            else
            {
                p.StartInfo.Arguments = @"/select," + Directory.GetCurrentDirectory();
                
            }
            p.Start();
        }
    }
}
