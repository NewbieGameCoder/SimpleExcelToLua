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
        const string columnTag = "|";

        int columnCount;
        string tableName;
        TableListNode<TableNode> tableColumnNodeList;
        List<string> tableCellDescList;
        List<string> tableCellNameList;
        List<string> tableCellTypeList;
        List<string> tableNestTagList;
        List<int> columnsInfo;
        List<List<CellInfo>> rowsInfo;
        HashSet<CellInfo> conflictKeySet;

        public ConfigInfo(string tableName, DataTable configTagble)
        {
            if (configTagble == null)
                return;

            this.tableName = tableName;
            columnCount = configTagble.Columns.Count;
            int rowCount = configTagble.Rows.Count;
            columnsInfo = new List<int>();
            conflictKeySet = new HashSet<CellInfo>();
            tableCellDescList = new List<string>(columnCount);
            tableCellNameList = new List<string>(columnCount);
            tableCellTypeList = new List<string>(columnCount);
            tableNestTagList = new List<string>(columnCount);
            rowsInfo = new List<List<CellInfo>>(rowCount - cellNameRowIndex);

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

                if (string.CompareOrdinal(nestTag, columnTag) == 0)
                {
                    columnsInfo.Add(i);
                }

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
                TabelHeirarchyIterator();
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

                int firstEleColumnIndex = columnsInfo[0];
                /// 如果默认没有采用列数据的话，就不再嵌套一层了
                if (tableColumnNodeList.Count == 1)
                {
                    tableColumnNodeList[0].NestInArray();
                    return tableColumnNodeList[0].Serialize(sb, indent);
                }
                else
                {
                    return tableColumnNodeList.Serialize(sb, indent);
                }
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
            string firstNodeType = tableCellTypeList[enumColumnIndex];
            endIndex = columnIndex;

            FilterEmptyColumn(ref enumColumnIndex, firstNodeType, firstNodeTag);
            if (!CacheNestTableType(enumColumnIndex, out bNodeCachedByList))
            {
                endIndex = columnIndex;
                return null;
            }

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
            else if (nestEleCount == 1 && Regex.IsMatch(firstNodeTag, dicTagRegex))
            {
                subNode = new CellInfo(null, null, "bool", "true");
                endIndex = enumColumnIndex;
                return subNode;
            }

            curRowNode.Add(subNode);
            ++cachedEleCount;
            ++enumColumnIndex;
            if (enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
            {
                endIndex = enumColumnIndex;
                if (string.CompareOrdinal(firstNodeTag, columnTag) == 0)
                    return subNode;

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
                    subNode.NestInTable(tableCellNameList[enumColumnIndex]);
                    enumColumnIndex = tempEndIndex;
                }
                else
                {
                    subNode.NestInTable(tableCellNameList[enumColumnIndex]);
                    ++enumColumnIndex;
                }

                /// 如果有明确的数字标明本“数组”类型的集合的元素的个数，则该集合里面的元素不显示key
                if (bNodeCachedByArray)
                {
                    subNode.NestInArray();
                }

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

        void TabelHeirarchyIterator()
        {
            int startColumnIndex = 0;
            string tempTag = tableNestTagList[startColumnIndex];

            FilterEmptyColumn(ref startColumnIndex, tableCellTypeList[startColumnIndex], tempTag);
            if (columnsInfo.Count == 0)
                columnsInfo.Add(startColumnIndex);

            tableColumnNodeList = new TableListNode<TableNode>(columnsInfo.Count);

            for (int i = 0, colLen = columnsInfo.Count; i < colLen; ++i)
            {
                startColumnIndex = columnsInfo[i];
                int endIndex = 0;
                tempTag = tableNestTagList[startColumnIndex];
                string columnName = tableCellNameList[startColumnIndex];
                int nestColumnEleCount = i < colLen - 1 ? columnsInfo[i + 1] - startColumnIndex : columnCount - startColumnIndex;

                if (string.CompareOrdinal(tempTag, columnTag) == 0)
                {
                    /// 默认不管该列配了多少个数据，只读一个数据，要读多行数据请在“行”中实现
                    tableColumnNodeList.Add(RowHeirarchyIterator(startColumnIndex, nestTagRowIndex + 1, 1, out endIndex));
                    ++startColumnIndex;
                    --nestColumnEleCount;
                    if (nestColumnEleCount <= 0)
                        continue;
                }

                bool bCacheByList = true;
                if (!CacheNestTableType(startColumnIndex, out bCacheByList))
                    return;

                if (bCacheByList)
                {
                    var curListNode = new TableListNode<TableNode>(rowsInfo.Count);
                    int itorIndex = startColumnIndex;

                    while (endIndex - startColumnIndex < nestColumnEleCount)
                    {
                        for (int j = nestTagRowIndex + 1, rowLen = rowsInfo.Count; j < rowLen; ++j)
                        {
                            curListNode.Add(RowHeirarchyIterator(itorIndex, j, nestColumnEleCount, out endIndex));
                        }
                        ++itorIndex;
                    }

                    curListNode.NestInTable(columnName);
                    tableColumnNodeList.Add(curListNode);
                }
                else
                {
                    var curDicNode = new TableDicNode<TableNode, TableNode>(rowsInfo.Count);
                    curDicNode.Desc = tableCellDescList[startColumnIndex];
                    int itorIndex = startColumnIndex;

                    while (endIndex - startColumnIndex < nestColumnEleCount)
                    {
                        for (int j = nestTagRowIndex + 1, rowLen = rowsInfo.Count; j < rowLen; ++j)
                        {
                            TableNode newNode = RowHeirarchyIterator(itorIndex, j, nestColumnEleCount, out endIndex);
                            AppendNewDicNode(curDicNode, rowsInfo[j][itorIndex], newNode);
                        }
                        ++itorIndex;
                    }

                    curDicNode.NestInTable(columnName);
                    tableColumnNodeList.Add(curDicNode);
                }
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

        bool CacheNestTableType(int checkColumnIndex, out bool bCacheByList)
        {
            bCacheByList = true;
            string tempTag = tableNestTagList[checkColumnIndex];
            if (Regex.IsMatch(tempTag, arrayTagRegex) || string.CompareOrdinal(tempTag, columnTag) == 0)
            {
                bCacheByList = true;
            }
            else if (Regex.IsMatch(tempTag, dicTagRegex))
            {
                bCacheByList = false;
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
        void NestInTable(string keyNameInNestedTable);
    }

    public class TableListNode<T> : List<T>, TableNode
        where T : TableNode
    {
        bool bNestInTable;
        bool bNestInArray;
        string strKeyNameInNestedTable;

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

        public void NestInTable(string keyNameInNestedTable)
        {
            bNestInTable = true;
            strKeyNameInNestedTable = keyNameInNestedTable;
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            int validContentLength = sb.Length;

            if (bNestInTable && !bNestInArray && !string.IsNullOrEmpty(strKeyNameInNestedTable))
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(strKeyNameInNestedTable);
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
        string strKeyNameInNestedTable;

        public string Desc { get; set; }

        public TableDicNode(int capacity)
            : base(capacity)
        {
        }

        public void NestInArray()
        {
            bNestInArray = true;
        }

        public void NestInTable(string keyNameInNestedTable)
        {
            bNestInTable = true;
            strKeyNameInNestedTable = keyNameInNestedTable;
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            int validContentLength = sb.Length;

            if (bNestInTable && !bNestInArray && !string.IsNullOrEmpty(strKeyNameInNestedTable))
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(strKeyNameInNestedTable);
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
        const string cellArrayTag = "|";

        bool bNestInArray;
        bool bNestInTable;
        bool bContainsArray;
        string cellTag;
        string originalData;
        string strKeyNameInNestedTable;
        Type cellType;

        public string Desc { get; set; }

        public CellInfo(string cellTag, string name, string type, string data)
        {
            strKeyNameInNestedTable = name;
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

            if (originalData.Contains(cellArrayTag))
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

        public void NestInTable(string keyNameInNestedTable)
        {
            bNestInTable = true;
            strKeyNameInNestedTable = !string.IsNullOrEmpty(keyNameInNestedTable) ? keyNameInNestedTable : strKeyNameInNestedTable;
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

            if (string.IsNullOrEmpty(originalData)
                || (string.IsNullOrEmpty(strKeyNameInNestedTable) 
                    && (cellType == null || !bNestInArray))
                /* && string.CompareOrdinal(cellTag, cellArrayTag) != 0*/)
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
                && (!bNestInArray && !string.IsNullOrEmpty(strKeyNameInNestedTable)))
            {
                keyNameSerialized = true;
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(strKeyNameInNestedTable);
                sb.Append(" = ");
            }

            string dataFormat = cellType == typeof(string) ? "\"{0}\"" : "{0}";
            if (originalData.Contains(cellArrayTag))
            {
                var tempStringArray = originalData.Split(cellArrayTag[0]);
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

        static object DefaultValueForType(Type targetType)
        {
            if (targetType == null)
                return "";

            return targetType.IsValueType ? Activator.CreateInstance(targetType) : "";
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
