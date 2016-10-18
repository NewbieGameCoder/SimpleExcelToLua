using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelToLua
{
    public class ExcelReader
    {
        static string strSheetName = "LuaTable$";

        /// <summary>
        /// 将指定Excel文件的内容读取到DataSet中
        /// </summary>
        public static DataTable ReadXlsxFile(string filePath)
        {
            // 检查文件是否存在且没被打开
            FileStream stream = null;
            try
            {
                stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("打开文件 {0} 失败", filePath));
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

            OleDbConnection conn = null;
            OleDbDataAdapter odda = null;
            DataSet ds = null;

            try
            {
                //此连接只能操作Excel2007之前(.xls)文件
                //string strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + FileFullPath + ";Extended Properties='Excel 8.0; HDR=NO; IMEX=1'";
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";

                conn = new OleDbConnection(strConn);
                conn.Open();

                ds = new DataSet();
                odda = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", strSheetName), conn);                    //("select * from [Sheet1$]", conn);
                odda.Fill(ds, strSheetName);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.ToString());
                return null;
            }
            finally
            {
                conn.Close();
                // 由于C#运行机制，即便因为表格中没有Sheet名为data的工作簿而return null，也会继续执行finally，而此时da为空，故需要进行判断处理
                if (odda != null)
                    odda.Dispose();
                conn.Dispose();
            }

            return ds.Tables[0];
        }
    }
}
