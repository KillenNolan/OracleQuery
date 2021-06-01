using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using WindowsFormsApp1.Common;

namespace Util
{
    //获得Excel默认解析的第一个文件表
    public class ExcelToSheet
    {
        private static IWorkbook hssfworkbook = null;
        private static ISheet sheet = null;

        public static ISheet getSheet(IFormFile file, int index)
        {

            //通过文件获得文件簿
            string fileName = file.FileName;
            string expandName = fileName.Substring(fileName.LastIndexOf("."));
            if (expandName.ToUpper().Contains(".XLSX"))
            {

                //07版本以上
                hssfworkbook = new XSSFWorkbook(file.OpenReadStream());
            }
            else if (expandName.ToUpper().Contains(".XLS"))
            {   //07版本以下
                hssfworkbook = new HSSFWorkbook(file.OpenReadStream());
            }
            else
            {
                throw new Exception("传入的不是Excel文件！");
            }

            //得到下标为index的工作表
            sheet = hssfworkbook.GetSheetAt(index);
            return sheet;
        }


        /// <summary>
        /// 从指定的文件路径下，指定的行数开始解析excel中的有效数据
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="rowNum">指定的行数</param>
        /// <param name="titleRow">指定的标题行</param>
        /// <returns></returns>
        public static DataTable getexcel(String fileName, int rowNum, int titleRow)
        {
            DataTable dt = new DataTable();
            try
            {
                IWorkbook workbook = null; //新建IWorkbook對象 
                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(fileStream); //xlsx數據讀入workbook 
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(fileStream); //xls數據讀入workbook 
                }
                ISheet sheet = workbook.GetSheetAt(0); //獲取第一個工作表 
                IRow row;// = sheet.GetRow(0); //新建當前工作表行數據 
                // MessageBox.Show(sheet.LastRowNum.ToString());
                row = sheet.GetRow(titleRow); //row讀入頭部
                if (row != null)
                {
                    for (int m = 0; m < row.LastCellNum; m++) //表頭 
                    {
                        string cellValue = row.GetCell(m).ToString(); //獲取i行j列數據 
                        //Console.WriteLine(cellValue);
                        dt.Columns.Add(cellValue);
                    }
                }
                for (int i = rowNum; i <= sheet.LastRowNum; i++) //對工作表每一行 
                {
                    System.Data.DataRow dr = dt.NewRow();
                    row = sheet.GetRow(i); //row讀入第i行數據 
                    if (row != null)
                    {
                        for (int j = 0; j < row.LastCellNum; j++) //對工作表每一列 
                        {
                            string cellValue = row.GetCell(j).ToString(); //獲取i行j列數據 
                            //Console.WriteLine(cellValue);
                            dr[j] = cellValue;
                        }
                    }
                    dt.Rows.Add(dr);
                }
                //Console.ReadLine();//這個有問題,讀不出來,反正它只是debug用的,所以取消它
                fileStream.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }






    }

    class Export
    {



        /// <summary>
        /// NPOI方式创建并写EXCEL文件 100W 162s
        /// </summary>
        /// <param name="xlFile">EXCEL文件</param>
        /// <param name="sheetName">EXCEL表名称</param>
        /// <param name="str">输出字符串数组</param>
        /// <param name="row0">输出起始行</param>
        /// <param name="col0">输出起始列</param>
        /// <param name="nRow">行数</param>
        /// <param name="nCol">列数</param>
        /// <returns>true=输出EXCEL文件成功, false=输出EXCEL文件失败</returns>
        public static bool WriteNewExcel(string xlFile, string sheetName, string[,] str, int row0, int col0, int nRow, int nCol)
        {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(sheetName);

            for (int i = 0; i < nRow - 1; i++)
            {
                XSSFRow row = (XSSFRow)sheet.CreateRow(row0 + i);
                for (int j = 0; j < nCol - 1; j++)
                {
                    row.CreateCell(col0 + j).SetCellValue(str[i, j]);
                }
            }
            try
            {
                using (FileStream fs = new FileStream(xlFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs);
                }

            }
            catch (Exception ex)
            {

                //throw;
            }

            return true;
        }

        // 将数据集转换到Excel： ConvertDataTableToExcel   ConvertDataGridViewToExcel
        // 目前支持的数据类型有:DataTable,二维数组,二维交错数组,DataGridView,ArrayList
        // 2010.01.03 采用NPOI类库，改善操作速度,便于扩展

        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="data">一维数组</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        public static bool ConvertToExcel<T>(T[] data, string xlsSaveFileName, string sheetName)
        {
            FileStream fs = new FileStream(xlsSaveFileName, FileMode.Create);
            try
            {
                HSSFWorkbook newBook = new HSSFWorkbook();

                HSSFSheet newSheet = (HSSFSheet)newBook.CreateSheet(sheetName);//新建工作簿
                HSSFRow newRow = (HSSFRow)newSheet.CreateRow(0);//创建行
                for (int i = 0; i < data.Length; i++)
                {
                    newSheet.GetRow(0).CreateCell(i).SetCellValue(Convert.ToDouble(data[i].ToString()));//写入数据
                }
                newBook.Write(fs);
                return true;
            }
            catch (Exception err)
            {
                throw new Exception("转换数据到Excel失败：" + err.Message);
            }
            finally
            {
                fs.Close();
            }
        }

        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="data">二维数组</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        public static bool ConvertToExcel<T>(T[,] data, string xlsSaveFileName, string sheetName)
        {
            FileStream fs = new FileStream(xlsSaveFileName, FileMode.Create);
            try
            {
                HSSFWorkbook newBook = new HSSFWorkbook();
                HSSFSheet newSheet = (HSSFSheet)newBook.CreateSheet(sheetName);//新建工作簿
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    HSSFRow newRow = (HSSFRow)newSheet.CreateRow(i);//创建行
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        newSheet.GetRow(i).CreateCell(j).SetCellValue(Convert.ToDouble(data[i, j].ToString()));//写入数据
                    }
                }
                newBook.Write(fs);
                return true;
            }
            catch (Exception err)
            {
                throw new Exception("转换数据到Excel失败：" + err.Message);
            }
            finally
            {
                fs.Close();
            }
        }

        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="data">交错数组</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        /// </summary>
        public static bool ConvertToExcel<T>(T[][] data, string xlsSaveFileName, string sheetName)
        {
            FileStream fs = new FileStream(xlsSaveFileName, FileMode.Create);
            try
            {
                HSSFWorkbook newBook = new HSSFWorkbook();
                HSSFSheet newSheet = (HSSFSheet)newBook.CreateSheet(sheetName);//新建工作簿
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    HSSFRow newRow = (HSSFRow)newSheet.CreateRow(i);//创建行
                    for (int j = 0; j < data[i].Length; j++)
                    {
                        newSheet.GetRow(i).CreateCell(j).SetCellValue(Convert.ToDouble(data[i][j].ToString()));//写入数据
                    }
                }
                newBook.Write(fs);
                return true;
            }
            catch (Exception err)
            {
                throw new Exception("转换数据到Excel失败：" + err.Message);
            }
            finally
            {
                fs.Close();
            }
        }


        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="dt">DataTable对象</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        public static bool ConvertToExcelKillen(System.Data.DataTable dt, string xlsSaveFileName, string sheetName)
        {
            FileStream fs = new FileStream(xlsSaveFileName, FileMode.Create);
            try
            {
                HSSFWorkbook newBook = new HSSFWorkbook();
                ISheet[] arrSheet = new ISheet[dt.Rows.Count / 65535];

                for (int k = 1; k <= arrSheet.Length; k++)
                {
                    //foreach (var newSheet in arrSheet)
                    //{
                    //    newSheet = (HSSFSheet)newBook.CreateSheet("sheet"+k);//新建工作簿
                    //    for (int i = 0; i < dt.Rows.Count; i++)
                    //    {
                    //        HSSFRow newRow = (HSSFRow)newSheet.CreateRow(i);//创建行
                    //        for (int j = 0; j < dt.Columns.Count; j++)
                    //        {
                    //            newSheet.GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());//写入数据
                    //        }
                    //    }
                    //    newBook.Write(fs);
                    //    k++;

                    //}
                    arrSheet[k] = (HSSFSheet)newBook.CreateSheet(sheetName);//新建工作簿
                    for (int i = 0; i < dt.AsEnumerable().Skip(0).Take(65535).ToList().Count; i++)
                    {
                        HSSFRow newRow = (HSSFRow)arrSheet[k].CreateRow(i);//创建行
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            arrSheet[k].GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());//写入数据
                        }
                    }
                    newBook.Write(fs);
                }

                return true;
            }
            catch (Exception err)
            {
                throw new Exception("转换数据到Excel失败：" + err.Message);
            }
            finally
            {
                fs.Close();
            }
        }

        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="dt">DataTable对象</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        public static bool ConvertToExcel(System.Data.DataTable dt, string xlsSaveFileName, string sheetName)
        {
            FileStream fs = new FileStream(xlsSaveFileName, FileMode.Create);
            try
            {
                HSSFWorkbook newBook = new HSSFWorkbook();

                HSSFSheet newSheet = (HSSFSheet)newBook.CreateSheet(sheetName);//新建工作簿
                for (int i = 0; i < 65536; i++)
                {
                    HSSFRow newRow = (HSSFRow)newSheet.CreateRow(i);//创建行
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        newSheet.GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());//写入数据
                    }
                }
                newBook.Write(fs);
                return true;
            }
            catch (Exception err)
            {
                throw new Exception("转换数据到Excel失败：" + err.Message);
            }
            finally
            {
                fs.Close();
            }
        }

        /// <summary>
        /// 将数据集导出到Excel文件
        /// </summary>
        /// <param name="dt">DataGridView对象</param>
        /// <param name="xlsSaveFileName">Excel文件名称</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>是否转换成功</returns>
        public static bool ConvertToExcel(System.Windows.Forms.DataGridView dgv, string xlsSaveFileName, string sheetName)
        {
            return ConvertToExcel((System.Data.DataTable)dgv.DataSource, xlsSaveFileName, sheetName);
        }
    }
    class ExcelExport
    {
        //#region 导出SourceGrid数据（最新版，批量快速输出）
        ///// <summary>
        ///// 导出SourceGrid数据
        ///// </summary>
        ///// <param name="grid">SourceGrid</param>
        ///// <param name="rowsStr">需要导出的行</param>
        ///// <param name="colsStr">需要导出的列</param>

        ////Excel导出的时候有两种软件插件可以使用（一种是office一种wps），因为各个插件的dll使用的方法不一样，因此要判断用户安装了哪个软件。
        //public static void NewExportSourceGridCell(DataTable grid, List<int> rowsStr, List<int> colsStr)
        //{

        //    //个人做的是政府项目，讲究国产化，在这里我先判断用户是否安装了wps。
        //    string excelType = "wps";
        //    Type type;
        //    type = Type.GetTypeFromProgID("ET.Application");//V8版本类型
        //    if (type == null)//没有安装V8版本
        //    {
        //        type = Type.GetTypeFromProgID("Ket.Application");//V9版本类型
        //        if (type == null)//没有安装V9版本
        //        {
        //            type = Type.GetTypeFromProgID("Kwps.Application");//V10版本类型
        //            if (type == null)//没有安装V10版本
        //            {
        //                type = Type.GetTypeFromProgID("EXCEL.Application");//MS EXCEL类型
        //                excelType = "office";
        //                if (type == null)
        //                {
        //                    //ModuleBaseUserControl.ShowError("检测到您的电脑上没有安装office或WSP软件，请先安装！");
        //                    return;//没有安装Office软件
        //                }
        //            }
        //        }
        //    }
        //    if (excelType == "wps")
        //    {
        //        WpsExcel(type, grid, rowsStr, colsStr);
        //    }
        //    else
        //    {
        //        OfficeExcel(type, grid, rowsStr, colsStr);
        //    }
        //}



        ////安装了wps
        //public static void WpsExcel(Type type, DataTable grid, List<int> rowsStr, List<int> colsStr)
        //{
        //    dynamic _app = Activator.CreateInstance(type);  //根据类型创建App实例
        //    dynamic _workbook;  //声明一个文件
        //    _workbook = _app.Workbooks.Add(Type.Missing); //创建一个Excel
        //    ET.Worksheet objSheet; //声明Excel中的页
        //    objSheet = _workbook.ActiveSheet;  //创建一个Excel
        //    ET.Range range;
        //    try
        //    {
        //        range = objSheet.get_Range("A1", Missing.Value);
        //        object[,] saRet = new object[rowsStr.Count, colsStr.Count];  //声明一个二维数组
        //        for (int iRow = 0; iRow < rowsStr.Count; iRow++)  //把sourceGrid中的数据组合成二维数组
        //        {
        //            int row = rowsStr[iRow];
        //            for (int iCol = 0; iCol < colsStr.Count; iCol++)
        //            {
        //                int col = colsStr[iCol];
        //                saRet[iRow, iCol] = grid[row, col].Value;
        //            }
        //        }
        //        range.set_Value(ET.ETRangeValueDataType.etRangeValueDefault, saRet);  //把组成的二维数组直接导入range
        //        _app.Visible = true;
        //        _app.UserControl = true;
        //    }
        //    catch (Exception theException)
        //    {
        //        String errorMessage;
        //        errorMessage = "Error: ";
        //        errorMessage = String.Concat(errorMessage, theException.Message);
        //        errorMessage = String.Concat(errorMessage, " Line: ");
        //        errorMessage = String.Concat(errorMessage, theException.Source);
        //        MessageBox.Show(errorMessage, "Error");
        //    }
        //}

        ////安装了office
        //public static void OfficeExcel(Type type, DataTable grid, List<int> rowsStr, List<int> colsStr)
        //{

        //    try
        //    {
        //        range = objSheet.get_Range("A1", Missing.Value);
        //        range = range.get_Resize(rowsStr.Count, colsStr.Count);
        //        object[,] saRet = new object[rowsStr.Count, colsStr.Count];
        //        for (int iRow = 0; iRow < rowsStr.Count; iRow++)
        //        {
        //            int row = rowsStr[iRow];
        //            for (int iCol = 0; iCol < colsStr.Count; iCol++)
        //            {
        //                int col = colsStr[iCol];
        //                saRet[iRow, iCol] = grid[row, col].Value;
        //            }
        //        }
        //        range.set_Value(Missing.Value, saRet);

        //    }
        //    catch (Exception theException)
        //    {
        //        String errorMessage;
        //        errorMessage = "Error: ";
        //        errorMessage = String.Concat(errorMessage, theException.Message);
        //        errorMessage = String.Concat(errorMessage, " Line: ");
        //        errorMessage = String.Concat(errorMessage, theException.Source);
        //        MessageBox.Show(errorMessage, "Error");
        //    }
        //}
        //#endregion


        private static List<string> ToList(DataTable dt)
        {
            List<string> key = new List<string>(100);

            foreach (var item in dt.Columns)
            {
                key.Add(item.ToString());
            }


            return key;
        }
        private static string ToModule(DataTable dt)
        {
            //StringBuilder sb = new StringBuilder(500);
            //foreach (var item in dt.Columns)
            //{
            //    sb.Append(item + $":null,");
            //}

            Dictionary<string, string> keyValues = new Dictionary<string, string>(100);

            foreach (var item in dt.Columns)
            {
                keyValues.Add(item.ToString(), "");
            }


            var module = JsonConvert.DeserializeObject("");
            return "";
        }
        public static void ExportXLS(DataTable dt)
        {


            //ToModule(dt);
            //ToList(dt);

            //StreamWriter writer = new StreamWriter("filePath", false);
            //writer.WriteLine();

            #region 填充数据
            string[,] dataArray = new string[1 + dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)//填写列名
            {
                dataArray[0, i] = dt.Columns[i].ColumnName;

                for (int j = 0; j < dt.Rows.Count; j++)//填入数据
                {
                    dataArray[j + 1, i] = dt.Rows[j][i].ToString();
                }
            }
            #endregion

            var shhetName = $"{System.Guid.NewGuid().ToString("N")}";
            var filePath = @"D:\\TEMP\\EXCEL\\" + shhetName + ".xlsx";
            //Util.Export.ConvertToExcel<T>(T[,] dataArray, filePath, shhetName);
            Util.Export.WriteNewExcel(filePath, shhetName, dataArray, 0, 0, dt.Rows.Count, dt.Columns.Count);


        }



        private static IWorkbook getXSSFWorkbook(string filePath)
        {

            //IWorkbook workbook = null; //新建IWorkbook對象 
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet("sheet1");
            using (FileStream fs = File.Create(filePath))
            {
                xssfworkbook.Write(fs);
            }

            return xssfworkbook;

        }

        /// <summary>
        /// 导出Excel 10W 14s
        /// </summary>
        /// <param name="dt">数据源</param>
        public static void Export(DataTable dt)
        {
            var filePath = @"D:\TEMP\EXCEL\" + $"{System.Guid.NewGuid().ToString("N")}.xlsx";
            var workbook = new SXSSFWorkbook((XSSFWorkbook)getXSSFWorkbook(filePath), 500);
            ISheet Isheet = workbook.GetSheetAt(0);
            //ISheet Isheet = workbook.CreateSheet(filePath);



            int rowIndex = 0;

            using (FileStream fs = File.Open(filePath, FileMode.Open))
            {


                StringBuilder sb = new StringBuilder(500);
                foreach (var item in dt.Columns)
                {
                    sb.Append(item + ",");
                }
                string[] fi = sb.ToString().Split(',');

                foreach (var item in dt.AsEnumerable().Skip(10).Take(65535))
                {
                    IRow row = Isheet.CreateRow(rowIndex);
                    //Type t = item.GetType();
                    //PropertyInfo[] fsi = dt.Columns;
                    //IEnumerable<DataRow> allButFirst5Contactss = dt.AsEnumerable().Skip(10).Take(5).ToList();
                    for (int i = 0; i < fi.Length - 1; i++)
                    {
                        var f = fi[i];



                        row.CreateCell(i).SetCellValue(item[i].ToString());
                        //row.CreateCell(i).SetCellValue(getString(f.GetValue(item)));
                    }
                    rowIndex++;
                }
                workbook.Write(fs);
                workbook.Dispose();

            }
        }



        private string DelQuota(string str)
        {
            string result = str;
            string[] strQuota = { "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "`", ";", "'", ",", ".", "/", ":", "/,", "<", ">", "?" };
            for (int i = 0; i < strQuota.Length; i++)
            {
                if (result.IndexOf(strQuota[i]) > -1)
                    result = result.Replace(strQuota[i], "");
            }
            return result;
        }

        /// <summary>
        ///  导出Excel  10W 2s
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="path"></param>
        public static void WriteExcel(DataSet ds, string path)
        {
            try
            {

                path = ExcelHelper.CreateFolder() + $"{System.Guid.NewGuid().ToString("N")}.xlsx";
                //path = Directory.GetCurrentDirectory();
                long totalCount = ds.Tables[0].Rows.Count;


                long rowRead = 0;
                float percent = 0;

                StreamWriter sw = new StreamWriter(path, true);
                //StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding("gb2312"));
                StringBuilder sb = new StringBuilder();
                for (int k = 0; k < ds.Tables[0].Columns.Count; k++)
                {
                    sb.Append(ds.Tables[0].Columns[k].ColumnName.ToString() + "\t");
                }
                sb.Append(Environment.NewLine);

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    //rowRead++;
                    //percent = ((float)(100 * rowRead)) / totalCount;
                    ////  Pbar.Maximum = (int)totalCount;
                    ////  Pbar.Value = (int)rowRead;
                    //textMsg.Text += "\r\n" + "正在写入[" + percent.ToString("0.00") + "%]...的数据";
                    //System.Windows.Forms.Application.DoEvents();

                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {

                        //sw.Write(DelQuota(dt1.Rows[i][j].ToString()));
                        //sb.Append(DelQuota(ds.Tables[0].Rows[i][j].ToString()) + "\t");
                        sb.Append(ds.Tables[0].Rows[i][j].ToString() + "\t");
                    }
                    sb.Append(Environment.NewLine);
                }
                sw.Write(sb.ToString());
                sw.Flush();
                sw.Close();
                //return writeMessage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //return ex.Message;
            }

        }


        /// <summary>
        /// datable文件寫入到excel.xlsx文件中
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="sheetName">文件名称</param>
        /// <returns></returns>
        public static IWorkbook TableExcel(DataTable dataTable, string sheetName)
        {
            IWorkbook wkBook = new XSSFWorkbook();
            //cell設置文字居中getJudgeRecordsList
            ICellStyle newCellStyle = wkBook.CreateCellStyle();
            newCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            IFont font = wkBook.CreateFont();
            font.FontName = "黑体";
            font.FontHeightInPoints = ((short)10);//设置字体大小
            newCellStyle.SetFont(font);


            //创建sheet
            ISheet sheet = wkBook.CreateSheet(sheetName);
            //写入列
            IRow ColumnNameRow = sheet.CreateRow(0);
            //设置行高
            ColumnNameRow.Height = 300;
            for (int colunmNameIndex = 0; colunmNameIndex < dataTable.Columns.Count; colunmNameIndex++)
            {
                ColumnNameRow.CreateCell(colunmNameIndex).SetCellValue(dataTable.Columns[colunmNameIndex].ColumnName);
            }

            //写入行
            int WriteRowCount = 1;
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                //sheet表创建新的一行
                IRow newRow = sheet.CreateRow(WriteRowCount);
                //设置行高
                newRow.Height = 300;
                for (int column = 0; column < dataTable.Columns.Count; column++)
                {

                    newRow.CreateCell(column).SetCellValue(dataTable.Rows[row][column].ToString());

                }

                WriteRowCount++;  //写入下一行
            }
            //遍历cell设置样式
            for (int i = 0; i < dataTable.Rows.Count + 1; i++)
            {
                IRow temprow = sheet.GetRow(i);
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    ICell tempcell = temprow.GetCell(j);
                    tempcell.CellStyle = newCellStyle;
                }

            }


            return wkBook;

        }


        /// <summary>  
        /// 将DataTable数据导出到Excel文件中(xlsx)  
        /// </summary>  
        /// <param name="dt">数据表</param>  
        /// <param name="filePath">文件名称</param>  
        public static string TableToExcelForXLSX(DataTable dt, string filePath)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet(filePath);
            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            string strFileName = filePath + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
            using (var fs = File.Create(strFileName))
            {
                xssfworkbook.Write(fs);
                //自动列宽
                for (int i = 0; i <= dt.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i, true);
                }
                return "OK," + strFileName;
            }
        }
    }


}
