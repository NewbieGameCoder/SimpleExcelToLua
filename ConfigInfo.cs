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
        const string listTagRegex = @"(?<=\[)(\d*)(?=\])";///match []、[number]
        const string limitLengthArrayTagRegex = @"(?<=\[)(\d+)(?=\])";///match [number]
        const string dicTagRegex = @"(?<=\{)(\d*)(?=\})";///match {}、{number}
        const string tableTagRegex = @"(?<=\[)(\d*)(?=\])|(?<=\{)(\d*)(?=\})|(?<=\|)(\d*)";/// match arrayTagRegex or dicTagRegex
        const string columnTagRegex = @"(?<=\|)(\d*)";

        int columnCount;
        int rowCount;
        string tableName;
        TableListNode<TableNode> tableColumnNodeList;
        List<string> tableCellDescList;
        List<string> tableCellNameList;
        List<string> tableCellTypeList;
        List<string> tableNestTagList;
        Dictionary<int, int> columnsInfo;
        List<List<CellInfo>> rowsInfo;
        HashSet<CellInfo> conflictKeySet;

        public ConfigInfo(string tableName, DataTable configTagble)
        {
            if (configTagble == null)
                return;

            this.tableName = tableName;
            columnCount = configTagble.Columns.Count;
            rowCount = configTagble.Rows.Count;
            columnsInfo = new Dictionary<int, int>();
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

                int validRowCount = 0;
                bool bColumnHeadDetected = false;
                /// 缓存列头的有效数据位
                if (Regex.IsMatch(nestTag, columnTagRegex)
                    && SerializeNestDataLength(nestTag, rowCount - nestTagRowIndex - 1, ref validRowCount))
                {
                    columnsInfo.Add(i, validRowCount);
                    bColumnHeadDetected = true;
                }

                for (int j = 0; j < rowCount; ++j)
                {
                    string data = configTagble.Rows[j][i].ToString();
                    CellInfo tempCellInfo = new CellInfo(bColumnHeadDetected, nestTag, cellName, cellType, data);
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
                bool bSerializedSuc = false;

                sb.Append("return ");
                if (tableColumnNodeList.Count == 1)
                {
                    tableColumnNodeList[0].NestInArray();
                    return tableColumnNodeList[0].Serialize(sb, indent);
                }
                else
                {
                    sb.Append("\n{\n");
                }

                ToLuaText.AppendIndent(sb, indent);
                for (int i = 0, len = tableColumnNodeList.Count; i < len; ++i)
                {
                    int tempValidContentLength = sb.Length;

                    ++indent;
                    if (tableColumnNodeList[i].Serialize(sb, indent))
                    {
                        bSerializedSuc = true;
                        sb.Append(",\n");
                    }
                    --indent;

                    if (bSerializedSuc)
                        ToLuaText.AppendIndent(sb, indent);
                    else
                        sb.Remove(tempValidContentLength, sb.Length - tempValidContentLength);
                }

                if (tableColumnNodeList.Count > 1)
                {
                    sb.Append("}");
                }

                return bSerializedSuc;
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("序列化表 {0} 发生错误\n{1}", tableName, e.ToString()));
                return false;
            }
        }

        TableNode RowHeirarchyIterator(int columnIndex, int rowIndex, int maxNestColumnCount, ref int endIndex)
        {
            var curRowNode = new TableListNode<TableNode>(columnCount - columnIndex);
            int enumColumnIndex = columnIndex;
            int cachedEleCount = 0;
            int nestEleCount = maxNestColumnCount;
            bool bNodeCachedByList = true;
            bool bNodeCachedByArray = false;
            string firstNodeTag = tableNestTagList[enumColumnIndex];

            if (!FilterEmptyColumn(ref enumColumnIndex) || !CacheNestTableType(enumColumnIndex, out bNodeCachedByList))
            {
                endIndex = ++enumColumnIndex;
                return null;
            }

            if (SerializeNestDataLength(firstNodeTag, maxNestColumnCount, ref nestEleCount))
            {
                if (Regex.IsMatch(firstNodeTag, limitLengthArrayTagRegex) && nestEleCount > 1)
                {
                    bNodeCachedByArray = true;
                }
                nestEleCount = Math.Min(maxNestColumnCount, nestEleCount);
            }
            else if (nestEleCount == -1)
            {
                endIndex = enumColumnIndex;
                return null;
            }

            TableNode subNode = rowsInfo[rowIndex][enumColumnIndex];
            if (bNodeCachedByArray)
                subNode.NestInArray();
            /// 带key的不再将key也加入队列
            if (!Regex.IsMatch(firstNodeTag, dicTagRegex))
            {
                subNode.NestInTable(tableCellNameList[enumColumnIndex]);
                curRowNode.Add(subNode);
            }
            ++cachedEleCount;
            ++enumColumnIndex;
            if (enumColumnIndex >= endIndex || enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
            {
                endIndex = enumColumnIndex;
                return curRowNode;
            }

            while (true)
            {
                string cellTag = tableNestTagList[enumColumnIndex];
                string cellType = tableCellTypeList[enumColumnIndex];
                subNode = rowsInfo[rowIndex][enumColumnIndex];

                int subNestEleCount = 1;
                if (SerializeNestDataLength(cellTag, maxNestColumnCount - 1, ref subNestEleCount))
                {
                    int tempEndIndex = endIndex;
                    if (subNestEleCount > 1)
                    {
                        subNode = RowHeirarchyIterator(enumColumnIndex, rowIndex, maxNestColumnCount - 1, ref tempEndIndex);
                    }
                    else
                    {
                        tempEndIndex = enumColumnIndex + 1;
                    }

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
                else if (subNestEleCount == -1)
                {
                    endIndex = ++enumColumnIndex;
                    return null;
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

                if (enumColumnIndex >= endIndex || enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
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
            int endIndex = columnCount - 1;

            FilterEmptyColumn(ref startColumnIndex);
            if (columnsInfo.Count == 0)
                columnsInfo.Add(startColumnIndex, rowCount - nestTagRowIndex - 1);

            tableColumnNodeList = new TableListNode<TableNode>(columnsInfo.Count);
            string tempTag = tableNestTagList[startColumnIndex];

            bool bItorEnd = false;
            var itor = columnsInfo.GetEnumerator();
            itor.MoveNext();
            while (!bItorEnd)
            {
                startColumnIndex = itor.Current.Key;
                tempTag = tableNestTagList[startColumnIndex];
                int itorRowCount = itor.Current.Value;
                int nestColumnEleCount = columnCount - startColumnIndex;
                string columnName = tableCellNameList[startColumnIndex];
                if (itor.MoveNext())
                {
                    nestColumnEleCount = itor.Current.Key - startColumnIndex;
                }
                else
                {
                    bItorEnd = true;
                }

                if (Regex.IsMatch(tempTag, columnTagRegex))
                {
                    /// 默认不管该列配了多少个数据，只读一个数据，要读多行数据请在“行”中实现
                    var subNode = rowsInfo[nestTagRowIndex + 1][startColumnIndex];
                    subNode.NestInTable(columnName);
                    tableColumnNodeList.Add(subNode);
                    ++startColumnIndex;
                    --nestColumnEleCount;
                    if (nestColumnEleCount <= 0)
                        continue;
                }

                bool bCacheByList = true;
                if (!CacheNestTableType(startColumnIndex, out bCacheByList))
                {
                    itor.Dispose();
                    return;
                }

                endIndex = nestColumnEleCount + startColumnIndex;
                int itorIndex = startColumnIndex;

                if (bCacheByList)
                {
                    var curListNode = new TableListNode<TableNode>(itorRowCount);

                    for (int j = nestTagRowIndex + 1, len = itorRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        curListNode.Add(RowHeirarchyIterator(itorIndex, j, nestColumnEleCount, ref endIndex));
                    }

                    tableColumnNodeList.Add(curListNode);
                }
                else
                {
                    var curDicNode = new TableDicNode<TableNode, TableNode>(itorRowCount);
                    curDicNode.Desc = tableCellDescList[startColumnIndex];

                    for (int j = nestTagRowIndex + 1, len = itorRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        TableNode newNode = RowHeirarchyIterator(itorIndex, j, nestColumnEleCount, ref endIndex);
                        AppendNewDicNode(curDicNode, rowsInfo[j][itorIndex], newNode);
                    }

                    tableColumnNodeList.Add(curDicNode);
                }
            }
            itor.Dispose();
        }

        bool FilterEmptyColumn(ref int startColumnIndex)
        {
            string cellType = tableNestTagList[startColumnIndex];
            string cellTag = tableNestTagList[startColumnIndex];

            while (string.IsNullOrEmpty(cellType) && string.IsNullOrEmpty(cellTag))
            {
                if (startColumnIndex >= columnCount)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("{0} 空列过多，导致序列化失败", tableName));
                    --startColumnIndex;
                    return false;
                }

                cellType = tableNestTagList[startColumnIndex];
                cellTag = tableNestTagList[startColumnIndex];

                ++startColumnIndex;
            }

            return true;
        }

        bool CacheNestTableType(int checkColumnIndex, out bool bCacheByList)
        {
            bCacheByList = true;
            string tempTag = tableNestTagList[checkColumnIndex];
            if (Regex.IsMatch(tempTag, listTagRegex))
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
            if (value == null)
                return;

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

        bool SerializeNestDataLength(string nestTag, int maxEleCount, ref int subNestEleCount)
        {
            var matchResult = Regex.Match(nestTag, tableTagRegex);
            if (matchResult.Success)
            {
                string strNestStructLen = matchResult.Value;

                if (string.IsNullOrEmpty(strNestStructLen))
                {
                    subNestEleCount = maxEleCount;
                }
                else if (!int.TryParse(strNestStructLen, out subNestEleCount) || subNestEleCount == 0
                    || (subNestEleCount == 1 && Regex.IsMatch(nestTag, dicTagRegex)))
                {
                    subNestEleCount = -1;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("表{0}对应的嵌套类型:{1}的集合长度序列化失败", tableName, nestTag));
                    return false;
                }

                return true;
            }
            else
            {
                return false;
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
        }

        public void NestInTable(string keyNameInNestedTable)
        {
            bNestInTable = true;
            strKeyNameInNestedTable = keyNameInNestedTable;
        }

        public bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            bool bSkipCurNest = Count == 1;
            bool bShowKeyName = bNestInTable && !bNestInArray && !bSkipCurNest;
            int validContentLength = sb.Length;

            if (bShowKeyName && !string.IsNullOrEmpty(strKeyNameInNestedTable))
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(strKeyNameInNestedTable);
                sb.Append(" = \n");
            }

            indent = bShowKeyName ? indent + 1 : indent;
            if (bSkipCurNest)
            {
                this[0].NestInArray();
                bSerializeSuc = this[0].Serialize(sb, indent);
            }
            else
                bSerializeSuc = ToLuaText.TransferList((List<T>)this, sb, indent);

            if (!bSerializeSuc)
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
        bool bColumnHead;
        string cellTag;
        string originalData;
        string strKeyNameInNestedTable;
        Type cellType;

        public string Desc { get; set; }

        public CellInfo(bool bColumnHead, string cellTag, string name, string type, string data)
        {
            strKeyNameInNestedTable = name;
            this.cellTag = cellTag;
            this.bColumnHead = bColumnHead;
            originalData = data;

            switch (type)
            {
                case "int":
                    cellType = typeof(int);
                    break;
                case "float":
                    cellType = typeof(float);
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

            if (string.IsNullOrEmpty(originalData) && cellType != null && cellTag == null)
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
            else if (cellType == typeof(float))
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
            bool bCellInValid = string.IsNullOrEmpty(originalData) || (string.IsNullOrEmpty(strKeyNameInNestedTable) && (cellType == null || !bNestInArray));
            bool bShowKeyName = (bNestInTable && !bNestInArray) || bColumnHead;

            if (bCellInValid && !bColumnHead)
                return bSerializeSuc;

            ToLuaText.AppendIndent(sb, indent);
            if (bShowKeyName && !string.IsNullOrEmpty(strKeyNameInNestedTable))
            {
                sb.Append(strKeyNameInNestedTable);
                sb.Append(" = ");
            }

            string dataFormat = cellType == typeof(string) ? "\"{0}\"" : "{0}";
            if (string.CompareOrdinal(cellTag, "[1]") == 0 || bContainsArray)
            {
                var tempStringArray = originalData.Split(cellArrayTag[0]);
                MethodInfo convertArrayDataGenericMethod = typeof(CellInfo).GetMethod("ConvertArrayData", BindingFlags.NonPublic | BindingFlags.Instance);
                var convertMethod = convertArrayDataGenericMethod.MakeGenericMethod(new Type[] { cellType });
                var transferMethod = ToLuaText.MakeGenericArrayTransferMethod(cellType);

                try
                {
                    var dataArray = convertMethod.Invoke(this, new object[] { tempStringArray });
                    sb.Append("{");
                    var arrayGetValueMethod = dataArray.GetType().GetMethod("GetValue", new Type[] { typeof(int) });
                    for (int i = 0; i < tempStringArray.Length; ++i)
                    {
                        sb.AppendFormat(dataFormat, arrayGetValueMethod.Invoke(dataArray, new object[] { i }).ToString().Replace("\n", @"\n").Replace("\"", @"\"""));
                        if (i < tempStringArray.Length - 1)
                            sb.Append(", ");
                    }
                    sb.Append("}");

                    bSerializeSuc = true;
                }
                catch (System.Exception e)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ConvertStringFillToArray ::src =" + originalData + "  " + e.Message);
                }
            }
            else
            {
                string tempStr = string.Format(dataFormat, originalData.Replace("\n", @"\n").Replace("\"", @"\"""));
                sb.Append(tempStr);
                bSerializeSuc = true;
            }

            string descFormatStr = null;
            if (!string.IsNullOrEmpty(originalData))
                descFormatStr = ", --{0}";
            else if (bColumnHead)
                descFormatStr = "--{0}";
            else
                bSerializeSuc = false;

            if (!string.IsNullOrEmpty(Desc))
                sb.AppendFormat(descFormatStr, Desc);

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
