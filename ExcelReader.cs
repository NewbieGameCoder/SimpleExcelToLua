using System;
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace ExcelToLua
{
    public class ExcelReader
    {
        static string strSheetName = "LuaTable";

        /// <summary>
        /// 将指定Excel文件的内容读取到DataSet中
        /// </summary>
        public static ExcelWorksheet ReadXlsxFile(string filePath)
        {
            // 检查文件是否存在且没被打开
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                ExcelPackage ep = new ExcelPackage(stream);
                ExcelWorksheets workdSheets = ep.Workbook.Worksheets;

                foreach (var sheet in workdSheets)
                {
                    if (string.Compare(sheet.Name, strSheetName, true) == 0)
                    {
                        return sheet;
                    }
                }

                return null;
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("打开文件 {0} 失败", Path.GetFileNameWithoutExtension(filePath)));
                return null;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                    stream.Dispose();
                }
            }
        }
    }
}
