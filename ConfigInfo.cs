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
                if (tableColumnNodeList.Count == 1)
                {
                    /// 处理的是整张表只有一个“|”标志的情况，即lua配置只有“一张”表。
                    /// "|"标志主要用来只用一个lua配置表，存储多种格式的lua配置表
                    tableColumnNodeList[0].NestInArray();
                    return tableColumnNodeList[0].Serialize(sb, indent);
                }

                return ToLuaText.TransferList(tableColumnNodeList, sb, indent);
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("序列化表 {0} 发生错误\n{1}", tableName, e.ToString()));
                return false;
            }
        }

        void TabelHeirarchyIterator()
        {
            bool bItorEnd = false;
            int startColumnIndex = 0;
            FilterEmptyColumn(ref startColumnIndex);
            if (columnsInfo.Count == 0)
                columnsInfo.Add(startColumnIndex, rowCount - nestTagRowIndex - 1);
            tableColumnNodeList = new TableListNode<TableNode>(columnsInfo.Count);
            Dictionary<int, int>.Enumerator itor = columnsInfo.GetEnumerator();
            itor.MoveNext();

            while (!bItorEnd)
            {
                startColumnIndex = itor.Current.Key;
                int partialRowCount = itor.Current.Value;
                int nestColumnEleCount = columnCount - startColumnIndex;
                if (itor.MoveNext())
                {
                    nestColumnEleCount = itor.Current.Key - startColumnIndex;
                }
                else bItorEnd = true;

                if (Regex.IsMatch(tableNestTagList[startColumnIndex], columnTagRegex))
                {
                    /// 默认不管该列配了多少个数据，只读一个数据，要读多行数据请在“行”中实现
                    var subNode = rowsInfo[nestTagRowIndex + 1][startColumnIndex];
                    subNode.NestInTable(tableCellNameList[startColumnIndex]);
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

                int endIndex = nestColumnEleCount + startColumnIndex;

                if (bCacheByList)
                {
                    var curListNode = new TableListNode<TableNode>(partialRowCount);

                    for (int j = nestTagRowIndex + 1, len = partialRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        curListNode.Add(RowHeirarchyIterator(startColumnIndex, j, nestColumnEleCount, 1, ref endIndex));
                    }

                    tableColumnNodeList.Add(curListNode);
                }
                else
                {
                    var curDicNode = new TableDicNode<TableNode, TableNode>(partialRowCount);
                    curDicNode.Desc = tableCellDescList[startColumnIndex];

                    for (int j = nestTagRowIndex + 1, len = partialRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        TableNode newNode = RowHeirarchyIterator(startColumnIndex, j, nestColumnEleCount, 1, ref endIndex);
                        AppendNewDicNode(curDicNode, rowsInfo[j][startColumnIndex], newNode);
                    }

                    tableColumnNodeList.Add(curDicNode);
                }
            }
            itor.Dispose();
        }

        TableNode RowHeirarchyIterator(int columnIndex, int rowIndex, int maxNestColumnCount, int rowIndent, ref int endIndex)
        {
            bool bNodeCachedByList = true;
            int enumColumnIndex = columnIndex;
            if (!FilterEmptyColumn(ref enumColumnIndex) || !CacheNestTableType(enumColumnIndex, out bNodeCachedByList))
            {
                endIndex = ++enumColumnIndex;
                return null;
            }

            bool firstEle = true;
            bool bNestInArray = false;
            int cachedEleCount = 0;
            int nestEleCount = maxNestColumnCount;
            var curRowNode = new TableListNode<TableNode>(columnCount - columnIndex);
            TableNode rowRootNode = curRowNode;

            if (!bNodeCachedByList && rowIndent > 1)
            {
                var tempDicNode = new TableDicNode<TableNode, TableNode>(columnCount - columnIndex);
                tempDicNode.Add(rowsInfo[rowIndex][enumColumnIndex], curRowNode);
                tempDicNode.Desc = tableCellDescList[enumColumnIndex];
                rowRootNode = tempDicNode;
            }

            while (true)
            {
                int subNestEleCount = 1;
                string cellTag = tableNestTagList[enumColumnIndex];
                string cellType = tableCellTypeList[enumColumnIndex];
                TableNode subNode = rowsInfo[rowIndex][enumColumnIndex];

                if (SerializeNestDataLength(cellTag, maxNestColumnCount - 1, ref subNestEleCount) || subNestEleCount != -1)
                {
                    int tempEndIndex = enumColumnIndex + 1;

                    if (firstEle)
                    {
                        nestEleCount = subNestEleCount;
                        if (Regex.IsMatch(cellTag, limitLengthArrayTagRegex) && nestEleCount > 1)
                        {
                            bNestInArray = true;
                        }

                        nestEleCount = Math.Min(maxNestColumnCount, nestEleCount);
                    }
                    else if (subNestEleCount > 1)
                    {
                        tempEndIndex = endIndex;
                        subNode = RowHeirarchyIterator(enumColumnIndex, rowIndex, maxNestColumnCount - 1, ++rowIndent, ref tempEndIndex);
                    }

                    subNode.Desc = tableCellDescList[enumColumnIndex];
                    subNode.NestInTable(tableCellNameList[enumColumnIndex]);
                    enumColumnIndex = tempEndIndex;
                    if (bNestInArray)
                    {
                        subNode.NestInArray();
                    }
                }
                else
                {
                    endIndex = ++enumColumnIndex;
                    return null;
                }

                ++cachedEleCount;
                if (bNodeCachedByList || (!bNodeCachedByList && !firstEle))
                {
                    curRowNode.Add(subNode);
                }
                firstEle = false;

                if (enumColumnIndex >= endIndex || enumColumnIndex >= columnCount || cachedEleCount >= nestEleCount)
                {
                    endIndex = enumColumnIndex;
                    break;
                }
            }

            return rowRootNode;
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
            else if (ToLuaText.TransferList((List<T>)this, sb, indent))
            {
                bSerializeSuc = true;
            }
            else if (bNestInArray)
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append("{}");
                bSerializeSuc = true;
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
            else if (bNestInArray)
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append("{}");
                bSerializeSuc = true;
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
        MethodInfo convertArrayDataGenericMethod;

        public string Desc { get; set; }

        public CellInfo(bool bColumnHead, string cellTag, string name, string type, string data)
        {
            originalData = data;
            this.cellTag = cellTag;
            this.bColumnHead = bColumnHead;
            strKeyNameInNestedTable = name;
            convertArrayDataGenericMethod = typeof(CellInfo).GetMethod("ConvertArrayData", BindingFlags.NonPublic | BindingFlags.Instance);

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

            if (originalData.Contains(cellArrayTag) || string.CompareOrdinal(cellTag, "[1]") == 0)
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

            if (bContainsArray)
            {
                var tempStringArray = originalData.Split(cellArrayTag[0]);
                var convertMethod = convertArrayDataGenericMethod.MakeGenericMethod(new Type[] { cellType });
                var transferMethod = ToLuaText.MakeGenericArrayTransferMethod(cellType);

                try
                {
                    var dataArray = convertMethod.Invoke(this, new object[] { tempStringArray });
                    transferMethod.Invoke(dataArray, new object[] { dataArray, sb, 0 });
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
                ToLuaText.AppendValue(cellType, originalData, sb);
                bSerializeSuc = true;
            }

            string descFormatStr = null;

            if (!string.IsNullOrEmpty(originalData))
                descFormatStr = ", --{0}";
            else if (bColumnHead)
                descFormatStr = "--{0}";
            else
                bSerializeSuc = false;

            /// 加上bColumnHead判断是避免没写Desc和有效数据（originalData为空）导致序列化出错
            if (!string.IsNullOrEmpty(Desc) || bColumnHead)
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
