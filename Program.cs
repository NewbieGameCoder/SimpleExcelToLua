using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;

namespace ExcelToLua
{
    class Program
    {
        static void FileDirectoryEnsurance(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        
        static void ClearDirectory(string filePath)
        {
            var existFiles = Directory.GetFiles(filePath, "*.*");
            for (int i = 0; i < existFiles.Length; ++i)
            {
                File.Delete(existFiles[i]);
            }
        }

        static void CopyNonExcelFiles(string excelFolderPath, string clientLuaFolderPath, string serverLuaFolderPath)
        {
            var copyingExtraFiles = Directory.GetFiles(excelFolderPath, "*.*", SearchOption.AllDirectories);
            for (int i = 0; i < copyingExtraFiles.Length; ++i)
            {
                string fileName = Path.GetFileName(copyingExtraFiles[i]);
                if (Path.GetExtension(copyingExtraFiles[i]) == ".xlsx")
                    continue;

                File.Copy(copyingExtraFiles[i], serverLuaFolderPath + "/" + fileName, true);
                File.Copy(copyingExtraFiles[i], clientLuaFolderPath + "/" + fileName, true);
            }
        }

        static void Main(string[] args)
        {
            string excelFolderPath = Path.GetFullPath("Excel");
            string serverLuaFolderPath = Path.GetFullPath("ServerLua");
            string clientLuaFolderPath = Path.GetFullPath("ClientLua");

            FileDirectoryEnsurance(excelFolderPath);
            FileDirectoryEnsurance(serverLuaFolderPath);
            FileDirectoryEnsurance(clientLuaFolderPath);
            ClearDirectory(serverLuaFolderPath);
            ClearDirectory(clientLuaFolderPath);
            //CopyNonExcelFiles(excelFolderPath, clientLuaFolderPath, serverLuaFolderPath);

            string[] excelFiles = Directory.GetFiles(excelFolderPath, "*.xlsx");
            if (excelFiles == null || excelFiles.Length == 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("没有找到任何的后缀为xlsx的Excel文件");
                Console.ReadKey();
                return;
            }

            List<ConfigInfo> luaConfigFileInfoList = new List<ConfigInfo>(excelFiles.Length);
            ExcelReader fileReader = new ExcelReader();

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < excelFiles.Length; ++i)
            {
                string fileName = Path.GetFileNameWithoutExtension(excelFiles[i]);
                var luaConfigFileInfo = new ConfigInfo(fileName, ExcelReader.ReadXlsxFile(excelFiles[i]));

                sb.Clear();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(string.Format("开始转换表 {0}", fileName));
                luaConfigFileInfo.Serialize(sb, 0);
                File.WriteAllText(serverLuaFolderPath + "/" + fileName + ".lua", sb.ToString(), new UTF8Encoding(false));
                File.WriteAllText(clientLuaFolderPath + "/" + fileName + ".lua", sb.ToString(), new UTF8Encoding(false));
            }

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("请查看具体的转换消息来判断是否全部转换成功");
            Console.ReadKey();
        }
    }
}
