using System;
using System.Data;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelToLua
{
    public class ConfigInfo : LuaSerialization
    {
        const int cellDescRowIndex = 0;
        const int cellNameRowIndex = 1;
        const int cellTypeRowIndex = 2;
        const int nestTagRowIndex = 3;
        const string arrayTagRegex = @"(?<=\[)(\d*)(?=\])";///match []、[number]
        const string dicTagRegex = @"(?<=\{)(\d*)(?=\})";///match {}、{number}
        const string tableTagRegex = @"(?<=\[)(\d*)(?=\])|(?<=\{)(\d*)(?=\})";/// match arrayTagRegex or dicTagRegex

        int columnCount;
        string tableName;
        TableNode tabelNodeRoot;
        List<string> tableCellDescList;
        List<string> tableCellNameList;
        List<string> tableCellTypeList;
        List<string> tableNestTagList;
        List<List<CellInfo>> rowsInfo;
        HashSet<CellInfo> conflictKeySet;

        public ConfigInfo(string tableName, DataTable configTagble)
        {
            if (configTagble == null)
                return;

            this.tableName = tableName;
            columnCount = configTagble.Columns.Count;
            int rowCount = configTagble.Rows.Count;
            rowsInfo = new List<List<CellInfo>>(rowCount - cellNameRowIndex);
            conflictKeySet = new HashSet<CellInfo>();
            tableCellDescList = new List<string>(columnCount);
            tableCellNameList = new List<string>(columnCount);
            tableCellTypeList = new List<string>(columnCount);
            tableNestTagList = new List<string>(columnCount);

            for (int i = 0; i < columnCount; ++i)
            {
                string cellDesc = configTagble.Rows[cellDescRowIndex][i].ToString();
                string cellName = configTagble.Rows[cellNameRowIndex][i].ToString().Trim();
                string cellType = configTagble.Rows[cellTypeRowIndex][i].ToString().Trim().ToLower();
                string nestTag = configTagble.Rows[nestTagRowIndex][i].ToString().Trim();

                tableCellDescList.Add(cellDesc);
                tableCellNameList.Add(cellName);
                tableCellTypeList.Add(cellType);
                tableNestTagList.Add(nestTag);

                for (int j = 0; j < rowCount; ++j)
                {
                    string data = configTagble.Rows[j][i].ToString();
                    CellInfo tempCellInfo = new CellInfo(nestTag, cellName, cellType, data);
                    tempCellInfo.Desc = cellDesc;

                    if (j >= rowsInfo.Count)
                        rowsInfo.Add(new List<CellInfo>(columnCount));

                    if (i >= rowsInfo[j].Count)
                        rowsInfo[j].Add(tempCellInfo);
                    else
                        Console.WriteLine(i);
                }
            }

            try
            {
                tabelNodeRoot = TabelHeirarchyIterator();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("序列化表 {0} 发生错误\n{1}", tableName, e.ToString()));
            }
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            try
            {
                sb.Append("return ");
                return tabelNodeRoot.Serialize(sb, indent);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("序列化表 {0} 发生错误\n{1}", tableName, e.ToString()));
                return false;
            }
        }

        TableNode RowHeirarchyIterator(int columnIndex, int rowIndex, int nestEleCount, out int endIndex)
        {
            var curRowNode = new TableListNode<TableNode>(columnCount - columnIndex);
            int enumColumnIndex = columnIndex;
            int cachedEleCount = 0;
            bool bNodeCachedByList = true;
            bool bNodeCachedByArray = false;
            string firstNodeTag = tableNestTagList[enumColumnIndex];
            endIndex = columnIndex;

            FilterEmptyColumn(ref enumColumnIndex, tableCellTypeList[enumColumnIndex], firstNodeTag);
            if (!CacheNestTableType(enumColumnIndex, out bNodeCachedByList))
                return null;

            var matchResult = Regex.Match(firstNodeTag, tableTagRegex);
            if (matchResult.Success)
            {
                string strNestStructLen = matchResult.Value;

                if (!string.IsNullOrEmpty(strNestStructLen))
                {
                    if (!int.TryParse(strNestStructLen, out nestEleCount) || nestEleCount == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(string.Format("{0} 第{1}列的集合长度序列化失败", tableName, enumColumnIndex));
                        endIndex = enumColumnIndex;
                        return null;
                    }
                    else if (Regex.IsMatch(firstNodeTag, arrayTagRegex))
                    {
                        bNodeCachedByArray = true;
                    }
                }
            }

            TableNode subNode = rowsInfo[rowIndex][enumColumnIndex];
            if (bNodeCachedByArray)
            {
                subNode.NestInArray();
            }
            else if (nestEleCount == 1)
            {
                subNode = new CellInfo(null, null, "bool", "true");
                return subNode;
            }
            curRowNode.Add(subNode);
            ++cachedEleCount;
            ++enumColumnIndex;
            if (enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
            {
                endIndex = enumColumnIndex;
                return curRowNode;
            }

            while (true)
            {
                string cellTag = tableNestTagList[enumColumnIndex];
                string cellType = tableCellTypeList[enumColumnIndex];
                matchResult = Regex.Match(cellTag, tableTagRegex);
                subNode = rowsInfo[rowIndex][enumColumnIndex];

                if (matchResult.Success)
                {
                    int subNestEleCount;
                    string strNestStructLen = matchResult.Value;

                    if (string.IsNullOrEmpty(strNestStructLen))
                    {
                        subNestEleCount = columnCount;
                    }
                    else if (!int.TryParse(strNestStructLen, out subNestEleCount) || subNestEleCount == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(string.Format("{0} 第{1}列的集合长度序列化失败", tableName, enumColumnIndex));
                        endIndex = enumColumnIndex;
                        return null;
                    }

                    int tempEndIndex = enumColumnIndex;
                    subNode = RowHeirarchyIterator(enumColumnIndex, rowIndex, subNestEleCount, out tempEndIndex);
                    if (Regex.IsMatch(cellTag, dicTagRegex))
                    {
                        var tempDicNode = new TableDicNode<CellInfo, TableNode>(1);
                        tempDicNode.Add(rowsInfo[rowIndex][enumColumnIndex], subNode);
                        subNode = tempDicNode;
                    }
                    subNode.Desc = tableCellDescList[enumColumnIndex];
                    enumColumnIndex = tempEndIndex;
                }
                else
                {
                    ++enumColumnIndex;
                }

                /// 如果有明确的数字标明本“数组”类型的集合的元素的个数，则该集合里面的元素不显示key
                if (bNodeCachedByArray)
                {
                    subNode.NestInArray();
                }
                subNode.NestInTable();

                curRowNode.Add(subNode);
                ++cachedEleCount;

                if (enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
                {
                    endIndex = enumColumnIndex;
                    break;
                }
            }

            return curRowNode;
        }

        TableNode TabelHeirarchyIterator()
        {
            bool bUseArray = true;
            int startColumnIndex = 0;

            FilterEmptyColumn(ref startColumnIndex, tableCellTypeList[startColumnIndex], tableNestTagList[startColumnIndex]);
            if (!CacheNestTableType(startColumnIndex, out bUseArray))
                return null;

            int endIndex = 0;
            if (bUseArray)
            {
                var curListNode = new TableListNode<TableNode>(rowsInfo.Count);

                for (int i = nestTagRowIndex + 1, len = rowsInfo.Count; i < len; ++i)
                {
                    curListNode.Add(RowHeirarchyIterator(startColumnIndex, i, columnCount, out endIndex));
                }

                return curListNode;
            }
            else
            {
                var curDicNode = new TableDicNode<TableNode, TableNode>(rowsInfo.Count);
                curDicNode.Desc = tableCellDescList[startColumnIndex];

                for (int i = nestTagRowIndex + 1, len = rowsInfo.Count; i < len; ++i)
                {
                    TableNode newNode = RowHeirarchyIterator(startColumnIndex, i, columnCount, out endIndex);
                    AppendNewDicNode(curDicNode, rowsInfo[i][startColumnIndex], newNode);
                }

                return curDicNode;
            }
        }

        void FilterEmptyColumn(ref int startColumnIndex, string cellType, string cellTag)
        {
            while (string.IsNullOrEmpty(cellType) && string.IsNullOrEmpty(cellTag))
            {
                if (startColumnIndex >= columnCount)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("{0} 空列过多，导致序列化失败", tableName));
                    --startColumnIndex;
                    return;
                }

                ++startColumnIndex;
                cellType = tableNestTagList[startColumnIndex];
                cellTag = tableNestTagList[startColumnIndex];
            }
        }

        bool CacheNestTableType(int checkColumnIndex, out bool bUseArray)
        {
            bUseArray = true;
            string tempTag = tableNestTagList[checkColumnIndex];
            if (Regex.IsMatch(tempTag, arrayTagRegex))
            {
                bUseArray = true;
            }
            else if (Regex.IsMatch(tempTag, dicTagRegex))
            {
                bUseArray = false;
                if (string.IsNullOrEmpty(tableCellTypeList[checkColumnIndex]))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("{0} 第{1}列的键值对型集合对应的key值无效", tableName, checkColumnIndex);
                    return false;
                }
            }

            return true;
        }

        void AppendNewDicNode(TableDicNode<TableNode, TableNode> dicCollection, CellInfo key, TableNode value)
        {
            TableNode existNode = null;
            if (!dicCollection.TryGetValue(key, out existNode))
            {
                dicCollection.Add(key, value);
            }
            else
            {
                TableListNode<TableNode> confictValueList = null;
                if (!conflictKeySet.Contains(key))
                {
                    confictValueList = new TableListNode<TableNode>(2);
                    confictValueList.Add(existNode);
                    dicCollection[key] = confictValueList;
                    conflictKeySet.Add(key);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("表{0}描述为{1}的列存在相同的键值", tableName, key.Desc));
                }
                else
                {
                    confictValueList = existNode as TableListNode<TableNode>;
                }

                confictValueList.Add(value);
            }
        }
    }

    public interface TableNode : LuaSerialization
    {
        string Desc { get; set; }

        void NestInArray();
        void NestInTable();
    }

    public class TableListNode<T> : List<T>, TableNode
        where T : TableNode
    {
        bool bNestInTable;
        bool bNestInArray;

        public string Desc { get; set; }

        public TableListNode(int capacity)
            : base(capacity)
        {
        }

        public void NestInArray()
        {
            bNestInArray = true;
            //var itor = GetEnumerator();
            //while (itor.MoveNext())
            //{
            //    if (itor.Current is CellInfo
            //        || typeof(TableDicNode<,>).IsAssignableFrom(itor.Current.GetType()))
            //    {
            //        itor.Current.NestInArray();
            //    }
            //}
            //itor.Dispose();
        }

        public void NestInTable()
        {
            bNestInTable = true;
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            int validContentLength = sb.Length;

            if (bNestInTable && !bNestInArray)
            {
                CellInfo firstEle = this[0] as CellInfo;
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(firstEle.Name);
                sb.Append(" = \n");
            }

            indent = bNestInTable && !bNestInArray ? indent + 1 : indent;
            if (ToLuaText.TransferList((List<T>)this, sb, indent))
            {
                bSerializeSuc = true;
            }
            else
            {
                if (bNestInArray)
                {
                    ToLuaText.AppendIndent(sb, indent);
                    sb.Append("{}");
                    bSerializeSuc = true;
                }
                else
                {
                    sb.Remove(validContentLength, sb.Length - validContentLength);
                }
            }

            return bSerializeSuc;
        }
    }

    public class TableDicNode<T, U> : Dictionary<T, U>, TableNode
        where T : TableNode
        where U : TableNode
    {
        bool bNestInArray;
        bool bNestInTable;

        public string Desc { get; set; }

        public TableDicNode(int capacity)
            : base(capacity)
        {
        }

        public void NestInArray()
        {
            bNestInArray = true;
        }

        public void NestInTable()
        {
            bNestInTable = true;
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            int validContentLength = sb.Length;

            if (bNestInTable && !bNestInArray)
            {
                var itor = GetEnumerator();
                itor.MoveNext();
                CellInfo firstEle = itor.Current.Key as CellInfo;
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(firstEle.Name);
                sb.Append(" = \n");
            }

            indent = bNestInTable && !bNestInArray ? indent + 1 : indent;
            ToLuaText.AppendIndent(sb, indent);
            sb.AppendFormat("--Key代表{0}\n", Desc);

            if (ToLuaText.TransferDic((Dictionary<T, U>)this, sb, indent))
            {
                bSerializeSuc = true;
            }
            else
            {
                if (bNestInArray)
                {
                    ToLuaText.AppendIndent(sb, indent);
                    sb.Append("{}");
                    bSerializeSuc = true;
                }
                else
                {
                    sb.Remove(validContentLength, sb.Length - validContentLength);
                }
            }

            return bSerializeSuc;
        }
    }

    public class CellInfo : TableNode
    {
        bool bNestInArray;
        bool bNestInTable;
        bool bContainsArray;
        string originalData;
        string keyName;
        string cellTag;
        Type cellType;

        public string Desc { get; set; }

        public string Name { get { return keyName; } }

        public CellInfo(string cellTag, string name, string type, string data)
        {
            this.keyName = name;
            this.cellTag = cellTag;
            originalData = data;

            switch (type)
            {
                case "int":
                    cellType = typeof(int);
                    break;
                case "float":
                    cellType = typeof(float);
                    break;
                case "bool":
                    cellType = typeof(bool);
                    break;
                case "string":
                    cellType = typeof(string);
                    break;
                default:
                    if (!string.IsNullOrEmpty(type))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(string.Format("未知的类型{0}", type));
                    }
                    break;
            }

            if (originalData.Contains("|"))
            {
                bContainsArray = true;
            }
        }

        public void NestInArray()
        {
            bNestInArray = true;

            if (string.IsNullOrEmpty(originalData) && cellType != null)
            {
                originalData = DefaultValueForType(cellType).ToString();
            }
        }

        public void NestInTable()
        {
            bNestInTable = true;
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(originalData))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Key值为空，序列化失败");
            }
            else if (cellType == typeof(float) || cellType == typeof(bool))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("Key值类型无法为 {0}", cellType));
            }

            string tempString = originalData;
            if (cellType == typeof(string))
            {
                tempString = string.Format("\"{0}\"", originalData);
            }

            return tempString;
        }

        public override bool Equals(object obj)
        {
            if (obj == this)
                return true;

            var other = obj as CellInfo;
            if (other == null)
                return false;

            return originalData == other.originalData && cellType == other.cellType;
        }

        public override int GetHashCode()
        {
            return originalData.GetHashCode() + cellType.GetHashCode();
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;

            if ((cellType == null && string.IsNullOrEmpty(keyName))
                || (string.IsNullOrEmpty(originalData)/* && string.IsNullOrEmpty(cellTag)*/))
            {
                return bSerializeSuc;
            }
            else if (cellType == typeof(bool))
            {
                indent = 1;
                if (!bNestInArray)
                {
                    /// delete ",\n"
                    sb.Remove(sb.Length - 2, 2);
                }
                else
                {
                    /// delete ", "
                    sb.Remove(sb.Length - 2, 2);
                }
            }

            bool keyNameSerialized = false;
            if ((!bContainsArray || bNestInTable)
                && (!bNestInArray && !string.IsNullOrEmpty(keyName)))
            {
                keyNameSerialized = true;
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(keyName);
                sb.Append(" = ");
            }

            string dataFormat = cellType == typeof(string) ? "\"{0}\"" : "{0}";
            if (originalData.Contains("|"))
            {
                var tempStringArray = originalData.Split('|');
                MethodInfo convertArrayDataGenericMethod = typeof(CellInfo).GetMethod("ConvertArrayData", BindingFlags.NonPublic | BindingFlags.Instance);
                var convertMethod = convertArrayDataGenericMethod.MakeGenericMethod(new Type[] { cellType });
                var transferMethod = ToLuaText.MakeGenericArrayTransferMethod(cellType);

                try
                {
                    var dataArray = convertMethod.Invoke(this, new object[] { tempStringArray });
                    if (!bNestInArray)
                    {
                        bool bSeirializeResult = (bool)transferMethod.Invoke(null, new object[] { dataArray, sb, keyNameSerialized ? 0 : indent });
                        if (bSeirializeResult)
                        {
                            bSerializeSuc = true;
                        }
                    }
                    else
                    {
                        ToLuaText.AppendIndent(sb, indent);
                        var arrayGetValueMethod = dataArray.GetType().GetMethod("GetValue", new Type[] { typeof(int) });
                        for (int i = 0; i < tempStringArray.Length; ++i)
                        {
                            sb.AppendFormat(dataFormat, arrayGetValueMethod.Invoke(dataArray, new object[] { i }));
                            if (i < tempStringArray.Length - 1)
                                sb.Append(", ");
                        }

                        bSerializeSuc = true;
                    }
                }
                catch (System.Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ConvertStringFillToArray ::src =" + originalData + "  " + e.Message);
                }
            }
            else
            {
                if (!keyNameSerialized)
                    ToLuaText.AppendIndent(sb, indent);

                string tempStr = string.Format(dataFormat, originalData).Replace("\n", @"\n");
                sb.Append(tempStr);
                bSerializeSuc = true;
            }

            if (!string.IsNullOrEmpty(Desc))
                sb.AppendFormat(", --{0}", Desc);

            return bSerializeSuc;
        }

        public static object DefaultValueForType(Type targetType)
        {
            if (targetType == null)
                return "nil";

            return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
        }

        T[] ConvertArrayData<T>(string[] stringData)
        {
            T[] data = new T[stringData.Length];
            for (int i = 0, len = stringData.Length; i < len; ++i)
            {
                data[i] = (T)System.Convert.ChangeType(stringData[i], cellType);
            }
            return data;
        }
    }
}
