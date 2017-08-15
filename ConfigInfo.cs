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
        const string columnTagRegex = @"(?<=\|)(\d*)";
        const string listTagRegex = @"(?<=\[)(\d*)(?=\])";///match []、[number]
        const string limitLengthArrayTagRegex = @"(?<=\[)(\d+)(?=\])";///match [number]
        const string dicTagRegex = @"(?<=\{)(\d*)(?=\})";///match {}、{number}
        const string tableTagRegex = @"(?<=\[)(\d*)(?=\])|(?<=\{)(\d*)(?=\})|(?<=\|)(\d*)";/// match arrayTagRegex or dicTagRegex

        int rowCount;
        int columnCount;
        string tableName;
        List<string> tableCellDescList;
        List<string> tableCellNameList;
        List<string> tableCellTypeList;
        List<string> tableNestTagList;
        List<List<CellInfo>> rowsInfo;
        Dictionary<int, int> columnsInfo;
        TableListNode tableColumnNodeList;

        public ConfigInfo(string tableName, DataTable configTagble)
        {
            if (configTagble == null)
                return;

            this.tableName = tableName;
            columnCount = configTagble.Columns.Count;
            rowCount = configTagble.Rows.Count;
            columnsInfo = new Dictionary<int, int>();
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
            bool bSerializeSec = false;
            try
            {
                sb.Append("return ");
                var nodeList = tableColumnNodeList.NodeContainer;
                if (nodeList.Count == 1)
                {
                    /// 处理的是整张表只有一个“|”标志的情况，即lua配置只有“一张”表。
                    /// "|"标志主要用来只用一个lua配置表，存储多种格式的lua配置表
                    nodeList[0].NestInArray();
                    bSerializeSec = nodeList[0].Serialize(sb, indent);
                }
                else
                {
                    bSerializeSec = ToLuaText.TransferList(nodeList, sb, indent);
                }

                /// 去掉最后一个","字符
                sb.Remove(sb.Length - 1, 1);
                return bSerializeSec;
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
            tableColumnNodeList = new TableListNode(columnsInfo.Count);
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
                    var curListNode = new TableListNode(partialRowCount);

                    for (int j = nestTagRowIndex + 1, len = partialRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        curListNode.Add(RowHeirarchyIterator(startColumnIndex, j, nestColumnEleCount, 1, ref endIndex));
                    }

                    tableColumnNodeList.Add(curListNode);
                }
                else
                {
                    var curDicNode = new TableDicNode(partialRowCount);
                    curDicNode.Desc = tableCellDescList[startColumnIndex];

                    for (int j = nestTagRowIndex + 1, len = partialRowCount + nestTagRowIndex + 1; j < len; ++j)
                    {
                        TableNode newNode = RowHeirarchyIterator(startColumnIndex, j, nestColumnEleCount, 1, ref endIndex);
                        curDicNode.Add(rowsInfo[j][startColumnIndex], newNode);
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
            var curRowNode = new TableListNode(columnCount - columnIndex);
            TableNode rowRootNode = curRowNode;

            if (!bNodeCachedByList && rowIndent > 1)
            {
                var tempDicNode = new TableDicNode(columnCount - columnIndex);
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

                if (SerializeNestDataLength(cellTag, maxNestColumnCount, ref subNestEleCount) || subNestEleCount != -1)
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
                if (bNestInArray)
                {
                    subNode.NestInArray();
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

    public abstract class TableNode : LuaSerialization
    {
        protected bool bNestInTable;
        protected bool bNestInArray;
        protected bool bSkipCurNest;
        protected string strKeyNameInNestedTable;

        public string Desc { get; set; }

        public abstract bool Serialize(StringBuilder sb, int indent);

        public virtual void NestInArray()
        {
            bNestInArray = true;
        }

        public virtual void NestInTable(string keyNameInNestedTable)
        {
            bNestInTable = true;
            if (!string.IsNullOrEmpty(keyNameInNestedTable))
                strKeyNameInNestedTable = keyNameInNestedTable;
        }

        public virtual void SkipCurNest()
        {
            bSkipCurNest = true;
        }
    }

    public abstract class CompositeNode : TableNode
    {
        public override bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            bool bShowKeyName = ShowKeyName();

            if (bShowKeyName)
            {
                ToLuaText.AppendIndent(sb, indent);
                sb.Append(strKeyNameInNestedTable);
                sb.Append(" = \n");
            }

            indent = bShowKeyName ? indent + 1 : indent;

            if (VariationalSerialize(sb, indent))
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

        protected abstract bool VariationalSerialize(StringBuilder sb, int indent);

        protected abstract bool ShowKeyName();
    }

    public class TableListNode : CompositeNode
    {
        List<TableNode> nodeContainer;

        public TableListNode(int capacity)
        {
            nodeContainer = new List<TableNode>(capacity);
        }

        public void Add(TableNode value)
        {
            nodeContainer.Add(value);
            bSkipCurNest = nodeContainer.Count == 1;
        }

        public List<TableNode> NodeContainer { get { return nodeContainer; } }

        protected override bool ShowKeyName()
        {
            return bNestInTable && !bNestInArray && !bSkipCurNest && !string.IsNullOrEmpty(strKeyNameInNestedTable);
        }

        protected override bool VariationalSerialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;

            if (bSkipCurNest)
            {
                for (int i = 0, len = nodeContainer.Count; i < len; ++i)
                {
                    if (nodeContainer.Count == 1)
                        nodeContainer[i].SkipCurNest();

                    if (nodeContainer[i].Serialize(sb, indent))
                    {
                        bSerializeSuc = true;
                        if (nodeContainer[i] is CellInfo)
                            sb.Append("\n");
                    }
                }
            }
            else if (ToLuaText.TransferList(nodeContainer, sb, indent))
            {
                bSerializeSuc = true;
            }

            return bSerializeSuc;
        }
    }

    public class TableDicNode : CompositeNode
    {
        Dictionary<TableNode, TableNode> nodeContainer;

        public TableDicNode(int capacity)
        {
            nodeContainer = new Dictionary<TableNode, TableNode>(capacity);
        }

        public void Add(TableNode key, TableNode value)
        {
            if (nodeContainer.ContainsKey(key))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(string.Format("描述为{0}的列存在相同的键值", key.Desc));
            }
            else
            {
                nodeContainer.Add(key, value);
            }

            bSkipCurNest = nodeContainer.Count == 1;
        }

        protected override bool ShowKeyName()
        {
            return bNestInTable && !bNestInArray && !string.IsNullOrEmpty(strKeyNameInNestedTable);
        }

        protected override bool VariationalSerialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            ToLuaText.AppendIndent(sb, indent);
            sb.AppendFormat("--Key代表{0}\n", Desc);

            if (ToLuaText.TransferDic(nodeContainer, sb, indent))
            {
                bSerializeSuc = true;
            }

            return bSerializeSuc;
        }
    }

    public class CellInfo : TableNode
    {
        const string cellArrayTag = "|";

        bool bColumnHead;
        bool bInvalidData;
        bool bContainsArray;
        string cellTag;
        string originalData;
        Type cellType;
        MethodInfo convertArrayDataGenericMethod;

        public CellInfo(bool bColumnHead, string cellTag, string name, string type, string data)
        {
            originalData = data;
            this.cellTag = cellTag;
            this.bColumnHead = bColumnHead;
            strKeyNameInNestedTable = name;
            bInvalidData = string.IsNullOrEmpty(originalData);
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

        public override void NestInArray()
        {
            bNestInArray = true;

            if (string.IsNullOrEmpty(originalData) && cellType != null && cellTag == null)
            {
                originalData = DefaultValueForType(cellType).ToString();
                bInvalidData = string.IsNullOrEmpty(originalData);
            }
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

        public override bool Serialize(StringBuilder sb, int indent)
        {
            bool bSerializeSuc = false;
            bool bInvalidKeyName = string.IsNullOrEmpty(strKeyNameInNestedTable);
            bool bCellInValid = bInvalidData || (bInvalidKeyName && (cellType == null || !bNestInArray));
            bool bShowKeyName = ((bNestInTable && !bNestInArray) || bColumnHead) && !bInvalidKeyName && !bSkipCurNest;

            if (bCellInValid && !bColumnHead)
                return bSerializeSuc;

            ToLuaText.AppendIndent(sb, indent);
            if (bShowKeyName)
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

            string descFormatStr = "--{0}";
            if (bColumnHead && bInvalidData)
                sb.Remove(sb.Length - 2, 2);
            else if (bInvalidData)
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
