using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace WindowsFormsApp1.Common
{
    internal class ExcelHelper
    {

        public static string WriteExcel(DataSet ds, string path)
        {
            try
            {
        
                path = CreateFolder()+"test.xlsx";
                //path = Directory.GetCurrentDirectory();
                long totalCount = ds.Tables[0].Rows.Count;
                string writeMessage = "共有: " + totalCount + "条数据";
                Thread.Sleep(1000);
                long rowRead = 0;
                float percent = 0;

                StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding("gb2312"));
                StringBuilder sb = new StringBuilder();
                for (int k = 0; k < ds.Tables[0].Columns.Count; k++)
                {
                    sb.Append(ds.Tables[0].Columns[k].ColumnName.ToString() + "\t");
                }
                sb.Append(Environment.NewLine);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    rowRead++;
                    percent = ((float)(100 * rowRead)) / totalCount;
                    //  Pbar.Maximum = (int)totalCount;
                    //  Pbar.Value = (int)rowRead;
                    writeMessage += "\r\n" + "正在写入[" + percent.ToString("0.00") + "%]...的数据";
                    System.Windows.Forms.Application.DoEvents();

                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {
                        sb.Append(ds.Tables[0].Rows[i][j].ToString() + "\t");
                    }
                    sb.Append(Environment.NewLine);
                }
                sw.Write(sb.ToString());
                sw.Flush();
                sw.Close();

                MessageBox.Show("已经生成指定的Excel文件");
                return writeMessage;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return ex.Message;
            }

        }

        public static string TableToExcelForXLS(DataTable dt, string file)
        {
            try
            {
                HSSFWorkbook hssfworkbook = new HSSFWorkbook();
                ISheet sheet = hssfworkbook.CreateSheet(file);
                IRow row = sheet.CreateRow(0);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Columns[j].ColumnName);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i < 65535)
                    {
                        IRow row2 = sheet.CreateRow(i + 1);
                        for (int k = 0; k < dt.Columns.Count; k++)
                        {
                            ICell cell2 = row2.CreateCell(k);
                            cell2.SetCellValue(dt.Rows[i][k].ToString());
                        }
                        continue;
                    }
                    Auto(i, hssfworkbook, dt);
                    break;
                }
                if (!Directory.Exists(ConfigurationSettings.AppSettings["PATH"]))
                {
                    Directory.CreateDirectory(ConfigurationSettings.AppSettings["PATH"]);
                }
                string strFileName = ConfigurationSettings.AppSettings["PATH"] + DateTime.Now.ToString("yyyy-MM-dd-24hhmm") + file + ".xls";
                using FileStream fs = File.Create(strFileName);
                hssfworkbook.Write(fs);
                return "OK," + strFileName;
            }
            catch (Exception ex)
            {
                return "Fail:" + ex.Message;
            }
        }

        public static void Auto(int i, HSSFWorkbook hssfworkbook, DataTable dt)
        {
            ISheet[] arrayA = new ISheet[10];
            int k = 1;
            while (k <= dt.Rows.Count / 65536)
            {
                arrayA[k] = hssfworkbook.CreateSheet("sheet" + k);
                while (i < dt.Rows.Count)
                {
                    IRow row1 = arrayA[k].CreateRow(i - (65536 * k - 1));
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                    if (i - (65536 * k - 1) >= 65535)
                    {
                        break;
                    }
                    i++;
                }
                k++;
                i++;
            }
        }

        public static void AutoXlsx(int i, XSSFWorkbook hssfworkbook, DataTable dt)
        {
            ISheet[] arrayA = new ISheet[10];
            int k = 1;
            while (k <= dt.Rows.Count / 1048576)
            {
                arrayA[k] = hssfworkbook.CreateSheet("sheet" + k);
                while (i < dt.Rows.Count)
                {
                    IRow row1 = arrayA[k].CreateRow(i - (1048576 * k - 1));
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                    if (i - (1048576 * k - 1) >= 1048575)
                    {
                        break;
                    }
                    i++;
                }
                k++;
                i++;
            }
        }


        public static string CreateFolder()
        {
            if (!Directory.Exists(ConfigurationSettings.AppSettings["PATH"]))
            {
                Directory.CreateDirectory(ConfigurationSettings.AppSettings["PATH"]);
            }
            return ConfigurationSettings.AppSettings["PATH"];
        }
        public static string TableToExcelForXLSX(DataTable dt, string file)
        {
            try
            {

                XSSFWorkbook xssfworkbook = new XSSFWorkbook();
                ISheet sheet = xssfworkbook.CreateSheet(file);
                IRow row = sheet.CreateRow(0);
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    ICell cell = row.CreateCell(k);
                    cell.SetCellValue(dt.Columns[k].ColumnName);
                }
                for (int j = 0; j < dt.Rows.Count; j++)
                {

                    if (j < 1048575)
                    {
                        IRow row2 = sheet.CreateRow(j + 1);
                        for (int l = 0; l < dt.Columns.Count; l++)
                        {
                            ICell cell2 = row2.CreateCell(l);
                            cell2.SetCellValue(dt.Rows[j][l].ToString());
                        }
                        continue;
                    }

                    AutoXlsx(j, xssfworkbook, dt);
                    break;
                }
                if (!Directory.Exists(ConfigurationSettings.AppSettings["PATH"]))
                {
                    Directory.CreateDirectory(ConfigurationSettings.AppSettings["PATH"]);
                }
                string strFileName = ConfigurationSettings.AppSettings["PATH"] + DateTime.Now.ToString("yyyy-MM-dd") + file + ".xlsx";
                using FileStream fs = File.Create(strFileName);
                xssfworkbook.Write(fs);
                for (int i = 0; i <= dt.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i, useMergedCells: true);
                }
                return "OK," + strFileName;
            }
            catch (Exception ex)
            {
                return "Fail:" + ex.Message;
            }
        }
    }
}
