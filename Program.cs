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

        static void Main(string[] args)
        {
            string excelFolderPath = Path.GetFullPath("Excel");
            string luaFolderPath = Path.GetFullPath("Lua");
            FileDirectoryEnsurance(excelFolderPath);
            FileDirectoryEnsurance(luaFolderPath);

            var existLuaFiles = Directory.GetFiles(luaFolderPath, "*.*");
            for (int i = 0; i < existLuaFiles.Length; ++i)
            {
                File.Delete(existLuaFiles[i]);
            }
            var copyingExtraFiles = Directory.GetFiles(excelFolderPath, "*.*", SearchOption.AllDirectories);
            for (int i = 0; i < copyingExtraFiles.Length; ++i)
            {
                string fileName = Path.GetFileName(copyingExtraFiles[i]);
                if (Path.GetExtension(copyingExtraFiles[i]) == ".xlsx")
                    continue;

                File.Copy(copyingExtraFiles[i], luaFolderPath + "/" + fileName, true);
            }

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
                File.WriteAllText(luaFolderPath + "/" + fileName + ".lua", sb.ToString(), new UTF8Encoding(false));
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(string.Format("转换表 {0} 完毕", fileName));
            }

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("请查看具体的转换消息来判断是否全部转换成功");
            Console.ReadKey();
        }
    }
}
