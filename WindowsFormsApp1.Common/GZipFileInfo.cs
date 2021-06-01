using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OracleQueryNew.WindowsFormsApp1.Common
{
    /// <summary>
    /// 要压缩的文件信息
    /// </summary>
    public class GZipFileInfo
    {
        /// <summary>
        /// 文件索引
        /// </summary>
        public int Index = 0;
        /// <summary>
        /// 文件相对路径，'/'
        /// </summary>
        public string RelativePath = null;
        public DateTime ModifiedDate;
        /// <summary>
        /// 文件内容长度
        /// </summary>
        public int Length = 0;
        public bool AddedToTempFile = false;
        public bool RestoreRequested = false;
        public bool Restored = false;
        /// <summary>
        /// 文件绝对路径,'\'
        /// </summary>
        public string LocalPath = null;
        public string Folder = null;

        public bool ParseFileInfo(string fileInfo)
        {
            bool success = false;
            try
            {
                if (!string.IsNullOrEmpty(fileInfo))
                {
                    // get the file information
                    string[] info = fileInfo.Split(',');
                    if (info != null && info.Length == 4)
                    {
                        this.Index = Convert.ToInt32(info[0]);
                        this.RelativePath = info[1].Replace("/", "\\");
                        this.ModifiedDate = Convert.ToDateTime(info[2]);
                        this.Length = Convert.ToInt32(info[3]);
                        success = true;
                    }
                }
            }
            catch
            {
                success = false;
            }
            return success;
        }
    }



    /// <summary>
    /// 文件压缩后的压缩包类
    /// </summary>
    public class GZipResult
    {
        /// <summary>
        /// 压缩包中包含的所有文件,包括子目录下的文件
        /// </summary>
        public GZipFileInfo[] Files = null;
        /// <summary>
        /// 要压缩的文件数
        /// </summary>
        public int FileCount = 0;
        public long TempFileSize = 0;
        public long ZipFileSize = 0;
        /// <summary>
        /// 压缩百分比
        /// </summary>
        public int CompressionPercent = 0;
        /// <summary>
        /// 临时文件
        /// </summary>
        public string TempFile = null;
        /// <summary>
        /// 压缩文件
        /// </summary>
        public string ZipFile = null;
        /// <summary>
        /// 是否删除临时文件
        /// </summary>
        public bool TempFileDeleted = false;
        public bool Errors = false;
    }
}
