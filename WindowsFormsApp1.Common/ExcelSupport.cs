using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure.Design;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OracleQueryNew.WindowsFormsApp1.Common
{


    //    class ExcelUtil
    //    {
    //        public static int PER_SHEET_LIMIT = 500000;

    //        public static SXSSFWorkbook getSXSSFWorkbookByPageThread(string[] title, string[][] values)
    //        {

    //            SXSSFWorkbook wb = new SXSSFWorkbook();

    //            int pageNum = values.Length / PER_SHEET_LIMIT;
    //            int lastCount = values.Length % PER_SHEET_LIMIT;

    //            if (values.Length > PER_SHEET_LIMIT)
    //            {

    //                ICellStyle style = wb.CreateCellStyle();
    //                int sheet = lastCount == 0 ? pageNum : pageNum + 1;
    //                CountDownLatch downLatch = new CountDownLatch(sheet);
    //                Executor executor = Executors.newFixedThreadPool(sheet);

    //                for (int c = 0; c <= pageNum; c++)
    //                {
    //                    int rowNum = PER_SHEET_LIMIT;
    //                    if (c == pageNum)
    //                    {
    //                        if (lastCount == 0)
    //                        {
    //                            continue;
    //                        }
    //                        rowNum = lastCount;
    //                    }
    //                    Sheet sheet = wb.createSheet("page" + c);
    //                    executor.execute(new PageTask(downLatch, sheet, title, style, rowNum, values));
    //                }
    //                try
    //                {
    //                    downLatch.await();
    //                }
    //                catch (InterruptedException e)
    //                {
    //                    e.printStackTrace();
    //                }
    //            }
    //            return wb;
    //        }

    //        private class final
    //        {
    //        }
    //    }
    ////    class PageTask ext Runnable
    //    {
    // private CountDownLatch countDownLatch;

    //    private Sheet sheet;
    //    private String[] title;
    //    private CellStyle style;
    //    private int b;
    //    private String[][] values;

    //    public PageTask(CountDownLatch countDownLatch, Sheet sheet, String[] title, CellStyle style, int b, String[][] values)
    //    {
    //        this.countDownLatch = countDownLatch;
    //        this.sheet = sheet;
    //        this.title = title;
    //        this.style = style;
    //        this.b = b;
    //        this.values = values;
    //    }

    //    @Override
    // public void run()
    //    {

    //        try
    //        {
    //            Row row = sheet.createRow(0);

    //            Cell cell = null;

    //            for (int i = 0; i < title.Length; i++)
    //            {
    //                cell = row.createCell(i);
    //                cell.setCellValue(title[i]);
    //                cell.setCellStyle(style);
    //            }

    //            for (int i = 0; i < b; i++)
    //            {
    //                row = sheet.createRow(i + 1);
    //                for (int j = 0; j < values[i].Length; j++)
    //                {
    //                    row.createCell(j).setCellValue(values[i][j]);
    //                }
    //            }
    //        }
    //        catch (Exception e)
    //        {
    //            e.printStackTrace();
    //        }
    //        finally
    //        {
    //            if (countDownLatch != null)
    //            {
    //                countDownLatch.countDown();
    //            }
    //        }
    //    }
    //}

    /// <summary>
    /// 电脑配置
    /// Intel(R) Xeon(R) CPU:E7-4860 v2@5.29HZ (TwoDimensArrToExcel:11%==>56%)(WriteExcel:11%==>56%);
    /// 记忆体:3.6/8.0(TwoDimensArrToExcel:44%==>95%)(WriteExcel:44%==>58%);
    /// 以太网络...
    /// 
    /// 测试表：Tabel
    /// 总数据：11427291
    /// 表结构：9个Columns(栏位)
    /// </summary>
    public class ExcelSupport
    {

        /// <summary>
        /// 将DataTable转换成  Two dimensional Arrays二维数组[,];Jagged array交错数组[][]
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="filePath">@"D:\\TEMP\\EXCEL\\" + shhetName + ".xlsx";</param>
        public static void DtToTwoDimensArr(DataTable dt, string filePath, string sheetName)
        {

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

            //通过NPOI组件写入Excel数据
            TwoDimensArrToExcel(dataArray, filePath, dt.Rows.Count, dt.Columns.Count);
        }



        /// <summary>
        /// NPOI方式创建并写EXCEL文件 1W 0.9s ;10W 9.2s;100W 162s
        /// </summary>
        /// <param name="xlFile">EXCEL文件 @"D:\\TEMP\\EXCEL\\" </param>
        /// <param name="sheetName">EXCEL表名称</param>
        /// <param name="str">输出字符串数组</param>
        /// <param name="row0">输出起始行</param>
        /// <param name="col0">输出起始列</param>
        /// <param name="nRow">行数</param>
        /// <param name="nCol">列数</param>
        /// <returns>true=输出EXCEL文件成功, false=输出EXCEL文件失败</returns>
        private static void TwoDimensArrToExcel(string[,] str, string filePath, int rowNum, int colNum)
        {
            {

            }

            try
            {
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                string sheetName = Path.GetFileNameWithoutExtension(filePath);
                //SXSSFWorkbook wb = ExcelUtil.getSXSSFWorkbookByPageThread(TITLE, content);

                IWorkbook workbook = null; //新建IWorkbook對象 
                                           //FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);


                if (filePath.IndexOf(".xlsx") > 0) // 2007版本 
                {
                    workbook = new XSSFWorkbook(); //xlsx數據讀入workbook 
                }
                else if (filePath.IndexOf(".xls") > 0) // 2003版本 
                {
                    workbook = new HSSFWorkbook(); //xls數據讀入workbook 
                }

                ISheet sheet = workbook.CreateSheet(sheetName); //创建第一個工作表 


                for (int i = 0; i < rowNum + 1; i++)//此处多了一行标题数据，因此+1
                {
                    IRow row = sheet.CreateRow(i);//创建一行数据
                    for (int j = 0; j < colNum; j++)
                    {
                        row.CreateCell(j).SetCellValue(str[i, j]);
                    }
                }
                stopwatch.Stop();
                var ss = "\r\n" + $"时间:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")},{stopwatch.Elapsed.TotalSeconds}";
                byte[] buffer = new byte[1024 * 1024 * 5];
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    xssfworkbook.Write(ms);
                //    buffer = ms.ToArray();
                //    ms.Close();
                //}
                //return buffer;

                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs);
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }

        }



        /// <summary>
        /// 删除特殊字符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
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
        ///  文件流 导出Excel 1W 0.07s; 10W 0.87s, 100W 6.05s 
        ///  数据内含特殊字符，可能导致文件无法打开，或者数据不准确
        ///  此方法只处理了数据内含有特殊字符 , 的
        /// </summary>
        /// <param name="ds">数据源</param>
        /// <param name="path">@"D:\\TEMP\\EXCEL\\"</param>
        public static void WriteExcel(DataSet ds, string path)
        {
            try
            {

                //path = path + $"{System.Guid.NewGuid().ToString("N")}IO.xlsx";
                //path = Directory.GetCurrentDirectory();
                long totalCount = ds.Tables[0].Rows.Count;


                long rowRead = 0;
                float percent = 0;

                //StreamWriter sw = new StreamWriter(path, true);
                StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding("UTF-8"));
                StringBuilder sb = new StringBuilder();
                for (int k = 0; k < ds.Tables[0].Columns.Count; k++)
                {
                    //sb.Append(ds.Tables[0].Columns[k].ColumnName.ToString() + "\t");
                    sb.Append(ds.Tables[0].Columns[k].ColumnName.ToString() + ",");
                    //sb.Append("\"" + ds.Tables[0].Columns[k].ColumnName.ToString() + "\"" + ",");
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
                        //sb.Append(ds.Tables[0].Rows[i][j].ToString() + "\t");
                        sb.Append(ds.Tables[0].Rows[i][j].ToString() + ",");
                        //sb.Append("\"" + ds.Tables[0].Rows[i][j].ToString() + "\"" + ",");
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
        /// 获取数据需要分的页数
        /// </summary>
        /// <param name="totalCount"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public static int[] GetPage(long totalCount, int size)
        {
            int page = 1;
            List<int> ll = new List<int>(10);

            int iSize = (int)(totalCount / size);
            int iRemainder = (int)(totalCount % size);
            while (totalCount - size >= 0)
            {
                ll.Add(page++);
                totalCount = totalCount - size;
            }


            //是否有余数
            if (iRemainder > 0)
            {
                ll.Add(page++);
            }

            //ll.Add((int)(totalCount - size * page));

            return ll.ToArray();
        }


        /// <summary>
        /// 处理禁止转义后的含有逗号的数据
        /// </summary>
        /// <param name="str">\tb'&$#\"z/wHP</param>
        /// <returns></returns>
        private static string DealWithQuota(string str)
        {
            str = str.Replace("\r", "").Replace("\n", "");
            return str.Contains(",") ? "\"" + str.Replace("\t", "") + "\"" : str;
            //var ss = str.Contains(",") ? "\"" + str.Replace("\t", "").Replace("\r", "").Replace("\n", "") + "\"" : str;

            //if (str.Contains(","))
            //{
            //    var temp = str;
            //    var tstr = "\"" + temp.Replace("\t", "").Replace("\r", "").Replace("\n", "") + "\"";
            //    str = tstr;
            //}
            //return str;
        }






        /// <summary>
        ///  文件流 导出Excel 1W 0.07s; 10W 0.87s, 100W 6.05s 
        ///  数据内含特殊字符，可能导致文件无法打开，或者数据不准确
        ///  Q:一串数字，前面为0易转成数字;A:加入\t禁止转义
        ///  Q:禁止转移后，字符串内含有逗号aa,523会导致数据错列;A:查找有逗号，去除禁止转义\t,加"aa,523"
        ///  Q:按上述处理后，还有个别数据换行,发现数据本身有换行，吐血;A:去除换行符('\r'是回車，'\n'是換行，前者使光標到行首，後者使光標下移一格。)
        /// </summary>
        /// <param name="ds">数据源</param>
        /// <param name="path">@"D:\\TEMP\\EXCEL\\"</param>
        public static void WriteExcelPage(DataTable dt, string path, int rows, int page)
        {
            try
            {
                IEnumerable<DataRow> allButFirst5Contactss = dt.AsEnumerable().Skip(page * rows).Take(rows).ToList();

                int iRows = allButFirst5Contactss.Count<DataRow>();

                //StreamWriter sw = new StreamWriter(path, true);
                StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding("UTF-8"));
                StringBuilder sb = new StringBuilder(1000);
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    sb.Append("\t" + dt.Columns[k].ColumnName.ToString() + ",");
                }

                sb.Append(Environment.NewLine);

                for (int i = 0; i < iRows; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        //sb.Append(dt.Rows[i][j].ToString() + ",");
                        sb.Append("\t" + DealWithQuota(dt.Rows[i][j].ToString()) + ",");
                        //var strt = dt.Rows[i]["USERPASSWORD"].ToString();
                    }
                    sb.Append(Environment.NewLine);
                }
                sw.Write(sb.ToString());
                sw.Flush();
                sw.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //return ex.Message;
            }

        }





        /// <summary>
        /// Compress
        /// </summary>
        /// <param name="files">Array of FileInfo objects to be included in the zip file</param>
        /// <param name="folders">Array of Folder string</param>
        /// <param name="lpBaseFolder">Base folder to use when creating relative paths for the files 
        /// stored in the zip file. For example, if lpBaseFolder is 'C:\zipTest\Files\', and there is a file 
        /// 'C:\zipTest\Files\folder1\sample.txt' in the 'files' array, the relative path for sample.txt 
        /// will be 'folder1/sample.txt'</param>
        /// <param name="lpDestFolder">Folder to write the zip file into</param>
        /// <param name="zipFileName">Name of the zip file to write</param>
        public static void Compress(FileInfo[] files, string[] folders, string lpBaseFolder, string lpDestFolder, string zipFileName)
        {
            //support compress folder
            IList<FileInfo> list = new List<FileInfo>();
            foreach (FileInfo li in files)
               list.Add(li); 

            foreach (string str in folders)
            {
                DirectoryInfo di = new DirectoryInfo(str);
                foreach (FileInfo info in di.GetFiles("*.*", SearchOption.AllDirectories))
                {
                    list.Add(info);
                }
            }

            ////return Compress(list.ToArray(), lpBaseFolder, lpDestFolder, zipFileName, true);
        }

        ///// <summary>
        ///// Compress
        ///// </summary>
        ///// <param name="files">Array of FileInfo objects to be included in the zip file</param>
        ///// <param name="lpBaseFolder">Base folder to use when creating relative paths for the files 
        ///// stored in the zip file. For example, if lpBaseFolder is 'C:\zipTest\Files\', and there is a file 
        ///// 'C:\zipTest\Files\folder1\sample.txt' in the 'files' array, the relative path for sample.txt 
        ///// will be 'folder1/sample.txt'</param>
        ///// <param name="lpDestFolder">Folder to write the zip file into</param>
        ///// <param name="zipFileName">Name of the zip file to write</param>
        ///// <param name="deleteTempFile">Boolean, true deleted the intermediate temp file, false leaves the temp file in lpDestFolder (for debugging)</param>
        //public static void Compress(FileInfo[] files, string lpBaseFolder, string lpDestFolder, string zipFileName, bool deleteTempFile)
        //{
        //    //GZipResult result = new GZipResult();

        //    try
        //    {
        //        if (!lpDestFolder.EndsWith("\\"))
        //        {
        //            lpDestFolder += "\\";
        //        }

        //        string lpTempFile = lpDestFolder + zipFileName + ".tmp";
        //        string lpZipFile = lpDestFolder + zipFileName;

        //        result.TempFile = lpTempFile;
        //        result.ZipFile = lpZipFile;

        //        if (files != null && files.Length > 0)
        //        {
        //            CreateTempFile(files, lpBaseFolder, lpTempFile, result);

        //            if (result.FileCount > 0)
        //            {
        //                CreateZipFile(lpTempFile, lpZipFile, result);
        //            }

        //            // delete the temp file
        //            if (deleteTempFile)
        //            {
        //                File.Delete(lpTempFile);
        //                //result.TempFileDeleted = true;
        //            }
        //        }
        //    }
        //    catch //(Exception ex4)
        //    {
        //        //result.Errors = true;
        //    }
        //    //return result;
        //}
        //private static void CreateZipFile(string lpSourceFile, string lpZipFile, GZipResult result)
        //{
        //    byte[] buffer;
        //    int count = 0;
        //    FileStream fsOut = null;
        //    FileStream fsIn = null;
        //    GZipStream gzip = null;

        //    // compress the file into the zip file
        //    try
        //    {
        //        fsOut = new FileStream(lpZipFile, FileMode.Create, FileAccess.Write, FileShare.None);
        //        gzip = new GZipStream(fsOut, CompressionMode.Compress, true);

        //        fsIn = new FileStream(lpSourceFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        //        buffer = new byte[fsIn.Length];
        //        count = fsIn.Read(buffer, 0, buffer.Length);
        //        fsIn.Close();
        //        fsIn = null;

        //        // compress to the zip file
        //        gzip.Write(buffer, 0, buffer.Length);

        //        result.ZipFileSize = fsOut.Length;
        //        result.CompressionPercent = GetCompressionPercent(result.TempFileSize, result.ZipFileSize);
        //    }
        //    catch //(Exception ex1)
        //    {
        //        result.Errors = true;
        //    }
        //    finally
        //    {
        //        if (gzip != null)
        //        {
        //            gzip.Close();
        //            gzip = null;
        //        }
        //        if (fsOut != null)
        //        {
        //            fsOut.Close();
        //            fsOut = null;
        //        }
        //        if (fsIn != null)
        //        {
        //            fsIn.Close();
        //            fsIn = null;
        //        }
        //    }
        //}

        //private static void CreateTempFile(FileInfo[] files, string lpBaseFolder, string lpTempFile, GZipResult result)
        //{
        //    byte[] buffer;
        //    int count = 0;
        //    byte[] header;
        //    string fileHeader = null;
        //    string fileModDate = null;
        //    string lpFolder = null;
        //    int fileIndex = 0;
        //    string lpSourceFile = null;
        //    string vpSourceFile = null;
        //    GZipFileInfo gzf = null;
        //    FileStream fsOut = null;
        //    FileStream fsIn = null;

        //    if (files != null && files.Length > 0)
        //    {
        //        try
        //        {
        //            result.Files = new GZipFileInfo[files.Length];

        //            // open the temp file for writing
        //            fsOut = new FileStream(lpTempFile, FileMode.Create, FileAccess.Write, FileShare.None);

        //            foreach (FileInfo fi in files)
        //            {
        //                lpFolder = fi.DirectoryName + "\\";
        //                try
        //                {
        //                    gzf = new GZipFileInfo();
        //                    gzf.Index = fileIndex;

        //                    // read the source file, get its virtual path within the source folder
        //                    lpSourceFile = fi.FullName;
        //                    gzf.LocalPath = lpSourceFile;
        //                    vpSourceFile = lpSourceFile.Replace(lpBaseFolder, string.Empty);
        //                    vpSourceFile = vpSourceFile.Replace("\\", "/");
        //                    gzf.RelativePath = vpSourceFile;

        //                    fsIn = new FileStream(lpSourceFile, FileMode.Open, FileAccess.Read, FileShare.Read);
        //                    buffer = new byte[fsIn.Length];
        //                    count = fsIn.Read(buffer, 0, buffer.Length);
        //                    fsIn.Close();
        //                    fsIn = null;

        //                    fileModDate = fi.LastWriteTimeUtc.ToString();
        //                    gzf.ModifiedDate = fi.LastWriteTimeUtc;
        //                    gzf.Length = buffer.Length;

        //                    fileHeader = fileIndex.ToString() + "," + vpSourceFile + "," + fileModDate + "," + buffer.Length.ToString() + "\n";
        //                    header = Encoding.Default.GetBytes(fileHeader);

        //                    fsOut.Write(header, 0, header.Length);
        //                    fsOut.Write(buffer, 0, buffer.Length);
        //                    fsOut.WriteByte(10); // linefeed

        //                    gzf.AddedToTempFile = true;

        //                    // update the result object
        //                    result.Files[fileIndex] = gzf;

        //                    // increment the fileIndex
        //                    fileIndex++;
        //                }
        //                catch //(Exception ex1)
        //                {
        //                    result.Errors = true;
        //                }
        //                finally
        //                {
        //                    if (fsIn != null)
        //                    {
        //                        fsIn.Close();
        //                        fsIn = null;
        //                    }
        //                }
        //                if (fsOut != null)
        //                {
        //                    result.TempFileSize = fsOut.Length;
        //                }
        //            }
        //        }
        //        catch //(Exception ex2)
        //        {
        //            result.Errors = true;
        //        }
        //        finally
        //        {
        //            if (fsOut != null)
        //            {
        //                fsOut.Close();
        //                fsOut = null;
        //            }
        //        }
        //    }

        //    result.FileCount = fileIndex;
        //}


        /// <summary>
        /// 压缩文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        public static void CompressFile(string filePath)
        {
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string desFolder = Path.GetDirectoryName(filePath);
                string saveName = fileName.Remove(fileName.Length - 2, 2);

                DirectoryInfo di = new DirectoryInfo(Path.GetDirectoryName(filePath));

                FileInfo[] files = di.GetFiles(saveName+ "*.*");

                //Compress四个参数分别是 "要压缩的文件集"“要压缩的目标目录”，“保存压缩文件的目录”，压缩文件名
                GZip.Compress(files, desFolder, desFolder, saveName + ".csv.gz");


                GZip.Decompress(desFolder, desFolder, saveName + ".gz");

                var dest = Path.GetDirectoryName(filePath) + "\\" + "test.csv";
                ////string[] n = { Path.GetDirectoryName(filePath) };
                ////Compress(afi, n, Path.GetDirectoryName(filePath), dest, "compress");
                //using (FileStream originalFileStream = new FileStream(filePath, FileMode.Open))
                //{
                //    using (FileStream compressedFileStream = File.OpenWrite($"{dest}.gz"))
                //    {
                //        using (GZipStream compressioonStream = new GZipStream(compressedFileStream, CompressionMode.Compress))
                //        {
                //            originalFileStream.CopyTo(compressioonStream);
                //            //GZFileName = $"{fileName}.gz";
                //        }
                //    }
                //}

                //using (FileStream originalFileStream = new FileStream(filePath, FileMode.Open))
                //{
                //    using (FileStream compressedFileStream = File.Create($"{dest}.gz"))
                //    {
                //        using (GZipStream compressioonStream = new GZipStream(compressedFileStream, CompressionMode.Compress))
                //        {
                //            originalFileStream.CopyTo(compressioonStream);
                //            //GZFileName = $"{fileName}.gz";
                //        }
                //    }
                //}

                //File.Delete(filePath);
            }
            catch (Exception ex)
            {


            }
        }

        //    public bool GzipCompress(string[] sourceFiles, string disPath)
        //    {
        //        if (!Directory.Exists(disPath)) MessageBox.Show("路径出错!", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1); if (!disPath.EndsWith("\\")) disPath += "\\"; string newName = "F" + Global.UserID + "-" + Guid.NewGuid().ToString() + ".zip"; strfileName = new FileInfo(Global.AccessoryPath + newName); 
        //        bool result = true; FileStream fs1 = null; FileStream fs2 = null; GZipStream zips = null;
        //        try
        //        {
        //            foreach (var sourceFile in sourceFiles)
        //            {
        //                if (sourceFile != null)
        //                {
        //                    int index = sourceFile.LastIndexOf("\\"); 
        //                    FFuJianName = FFuJianName + Global.UserID + "*" + sourceFile.Substring(index + 1) + "*" + newName + "*/"; 
        //                    fs1 = new FileStream(sourceFile, FileMode.Open, FileAccess.Read); 
        //                    fs2 = new FileStream(disPath + newName, FileMode.OpenOrCreate, FileAccess.Write);
        //                    zips = new GZipStream(fs2, CompressionMode.Compress); 
        //                    byte[] tempb = new byte[fs1.Length]; 
        //                    fs1.Read(tempb, 0, (int)fs1.Length);
        //                    byte[] exb = System.Text.Encoding.Unicode.GetBytes(Path.GetFileName(sourceFile));
        //                    byte[] lastb = new byte[fs1.Length + exb.Length + 1]; 
        //                    lastb[0] = Convert.ToByte(exb.Length); exb.CopyTo(lastb, 1); 
        //                    tempb.CopyTo(lastb, exb.Length + 1); zips.Write(lastb, 0, lastb.Length
        //}

    }
}
