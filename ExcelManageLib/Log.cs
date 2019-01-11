using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelManageLib.DataModel;

namespace ExcelManageLib.Common
{
    public static class Log
    {
        public static void  Write(string errorMes) //, string name, string fileName
        {
            if (ToolConfiguration.isWriteLog)
            {
                string logPath = String.Empty;
                if (String.IsNullOrEmpty(ToolConfiguration.LogPath))
                {
                    logPath = @"C:\log\";
                }
                else
                {
                    logPath = ToolConfiguration.LogPath;
                }

                var directoryPath = logPath;
                var Today = string.Format("{0:yyyyMMdd}", DateTime.Now);
                var fileName = "Log_" + Today + ".txt";
                var fullfilePath = directoryPath + fileName;
                try
                {
                    if (Directory.Exists(directoryPath))//判断是否存在
                    {
                        if (!File.Exists(fullfilePath))
                        {
                            File.Create(fullfilePath).Dispose();//创建文件  
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(directoryPath);//创建新路径
                        File.Create(fullfilePath).Dispose();//创建文件  

                    }
                    using (StreamWriter tw = File.AppendText(fullfilePath))
                    {
                        errorMes = DateTime.Now.ToString() + "  " + errorMes;
                        tw.WriteLine(errorMes);
                        tw.Close();
                    }
                }
                catch (Exception ex)
                {

                }
            }

        }
    }
}
