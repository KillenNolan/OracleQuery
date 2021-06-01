using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using OracleQueryNew.WindowsFormsApp1.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Common;

namespace WindowsFormsApp1
{
    public class SqlQuery : Form
    {
        private string ConnectionString = "";

        private IContainer components = null;

        private TextBox textSql;

        private Button button1;

        private ComboBox comlist;

        private TextBox textBox1;

        private TextBox textMsg;

        public SqlQuery()
        {
            InitializeComponent();
            GetDbName();
        }

        public void GetDbName()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            int a = ConfigurationSettings.AppSettings.Count;
            for (int i = 0; i < a && !(ConfigurationSettings.AppSettings.Keys[i].ToString() == "PATH"); i++)
            {
                comlist.Items.Add(ConfigurationSettings.AppSettings.Keys[i].ToString());
            }
            ConnectionString = ConfigurationSettings.AppSettings[comlist.Text];
        }

        private void comlist_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConnectionString = ConfigurationSettings.AppSettings[comlist.Text];
        }

        private async void button1_Click(object sender, EventArgs e)
        {

            textMsg.Text += "\r\n" + "执行中";
            textMsg.ForeColor = Color.Green;
            if (comlist.Text == "" || textSql.Text == "")
            {
                textMsg.Text += "廠區或者SQL不能為空";
                textMsg.ForeColor = Color.Red;
                return;
            }

            textMsg.Text += "\r\n" + "正在查询数据，请稍等。。。";


            DataTable dt = new DataTable();
            OracleHelper oracleHelper = new OracleHelper();
            dt = oracleHelper.ExcuteSqlReturnDataTable(textSql.Text, ConnectionString);

            long totalCount = dt.Rows.Count;
            textMsg.Text += "\r\n" + "共有: " + totalCount + "条数据,每次查询10W条";


            //int page = (int)(totalCount / 100);
            int iRows = 1500;
            var page = ExcelSupport.GetPage(totalCount, iRows);
            var docPath = ExcelHelper.CreateFolder();
            var shhetName = $"{System.Guid.NewGuid().ToString("N")}";
            var msg = "";

            // 将计算操作放到一个 Task<string> 中去，新开线程
            textMsg.Text += await Task.Run(() =>
            {
                for (int k = 0; k < page.Length; k++)
                {
                    var filePath = docPath + shhetName + $"_{page[k]}";
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    msg += "\r\n" + $"线程:{Thread.CurrentThread.ManagedThreadId.ToString()}开始执行;开始时间:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";



                    ExcelSupport.WriteExcelPage(dt, filePath + ".csv", iRows, k);
                    ExcelSupport.CompressFile(filePath + ".csv");

                    //Compress三个参数分别是“要压缩的目标目录”，“保存压缩文件的目录”，压缩文件名
                    //GZip.Compress(filePath, filePath, shhetName);
                    //ExcelSupport.DtToTwoDimensArr(dt, filePath + ".xlsx", "");


                    stopwatch.Stop();
                    msg += "\r\n" + $"线程:{Thread.CurrentThread.ManagedThreadId.ToString()}执行完成;结束时间:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}耗时:{stopwatch.Elapsed.TotalSeconds}秒";

                    //this.textMsg.Focus();//获取焦点
                    //this.textMsg.Select(this.textMsg.TextLength, 0);//光标定位到文本最后
                    //this.textMsg.ScrollToCaret();//滚动到光标处
                    //MessageBox.Show("已经生成指定的Excel文件");


                }

                //if (dt.Rows.Count > 0)
                //{

                //    if (dt.Columns[0].ColumnName == "MSG")
                //    {
                //        textMsg.Text += dt.Rows[0][0].ToString();
                //        textMsg.ForeColor = Color.Red;
                //    }
                //    else
                //    {

                //        //DataSet ds;  //ds已经读取到了数据
                //        //DataTable dt1 = ds.Tables[0];  //每次能读取一张表
                //        textMsg.Text += "\r\n" + $"ds时间:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";
                //        DataSet ds = new DataSet();
                //        ds.Tables.Add(dt);

                //        IEnumerable<DataRow> allButFirst5Contactss = dt.AsEnumerable().Skip(0).Take(dt.Rows.Count).ToList();


                //        var shhetName = $"{System.Guid.NewGuid().ToString("N")}";
                //        var filePath = ExcelHelper.CreateFolder();
                //        Stopwatch stopwatch = new Stopwatch();
                //        stopwatch.Start();
                //        textMsg.Text += "\r\n" + $"开始时间:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";


                //        //Export(dt);


                //        //WriteExcel(ds, ExcelHelper.CreateFolder());
                //        //Util.Export.ConvertToExcel(dt, filePath + shhetName, "sheet1");

                //        //Util.Export.ConvertToExcel<T>(allButFirst5Contactss, filePath + shhetName, "sheet1");

                //        //ExcelHelper.TableToExcelForXLS(dt, "sheet1");


                //        //Util.ExcelExport.ExportXLS(dt);



                //        //ExcelSupport.DtToTwoDimensArr(dt, filePath, shhetName);
                //        ExcelSupport.WriteExcel(ds, filePath + shhetName + ".xlsx");

                //        ExcelSupport.CompressFile(filePath + shhetName + ".xlsx");

                //        stopwatch.Stop();
                //        textMsg.Text += "\r\n" + $"结束时间:{DateTime.Today.ToString("yyyy/MM/dd HH:mm:ss")}耗时:{stopwatch.Elapsed.TotalSeconds}秒";
                //        this.textMsg.Focus();//获取焦点
                //        this.textMsg.Select(this.textMsg.TextLength, 0);//光标定位到文本最后
                //        this.textMsg.ScrollToCaret();//滚动到光标处
                //        MessageBox.Show("已经生成指定的Excel文件");
                //    }
                //}
                //else
                //{
                //    textMsg.Text += "NO DATA";
                //    textMsg.ForeColor = Color.Red;
                //}
                return msg;
            });

            //textMsg.Text += await doSomething();

            this.textMsg.Focus();//获取焦点
            this.textMsg.Select(this.textMsg.TextLength, 0);//光标定位到文本最后
            this.textMsg.ScrollToCaret();//滚动到光标处
        }


        //private Task<string> doSomething()
        //{  // 使用 lambda 表达式定义计算和返回工作

        //    textMsg.Text += "\r\n" + "正在查询数据，请稍等。。。";


        //    DataTable dt = new DataTable();
        //    OracleHelper oracleHelper = new OracleHelper();
        //    dt = oracleHelper.ExcuteSqlReturnDataTable(textSql.Text, ConnectionString);

        //    long totalCount = dt.Rows.Count;
        //    textMsg.Text += "\r\n" + "共有: " + totalCount + "条数据";


        //    //int page = (int)(totalCount / 100);
        //    int iRows = 100000;
        //    var page = ExcelSupport.GetPage(totalCount,iRows );
        //    var shhetName = $"{System.Guid.NewGuid().ToString("N")}";
        //    var msg = "";



        //    return t;
        //}

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.textSql = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.comlist = new System.Windows.Forms.ComboBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textMsg = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textSql
            // 
            this.textSql.Location = new System.Drawing.Point(3, 68);
            this.textSql.Multiline = true;
            this.textSql.Name = "textSql";
            this.textSql.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textSql.Size = new System.Drawing.Size(799, 277);
            this.textSql.TabIndex = 0;
            this.textSql.Text = "請輸入查詢語句";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(588, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 53);
            this.button1.TabIndex = 1;
            this.button1.Text = "Query";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comlist
            // 
            this.comlist.FormattingEnabled = true;
            this.comlist.Location = new System.Drawing.Point(118, 26);
            this.comlist.Name = "comlist";
            this.comlist.Size = new System.Drawing.Size(121, 20);
            this.comlist.TabIndex = 2;
            this.comlist.SelectedIndexChanged += new System.EventHandler(this.comlist_SelectedIndexChanged);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.Location = new System.Drawing.Point(76, 26);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(36, 22);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "廠區:";
            // 
            // textMsg
            // 
            this.textMsg.Location = new System.Drawing.Point(3, 351);
            this.textMsg.Multiline = true;
            this.textMsg.Name = "textMsg";
            this.textMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textMsg.Size = new System.Drawing.Size(785, 87);
            this.textMsg.TabIndex = 4;
            // 
            // SqlQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textMsg);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.comlist);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textSql);
            this.Name = "SqlQuery";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
