using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

public interface LuaSerialization
{
    bool Serialize(StringBuilder sb, int indent);
    //void UnSerilize(IntPtr L);
}

public static class ToLuaText
{
    static Type listTypeDefinition = typeof(List<>);
    static Type dictionaryTypeDefinition = typeof(Dictionary<,>);
    static MethodInfo listTransferGenericMethod;
    static MethodInfo arrayTransferGenericMethod;
    static MethodInfo dictionaryTransferGenericMethod;

    static ToLuaText()
    {
        var classType = typeof(ToLuaText);
        listTransferGenericMethod = classType.GetMethod("TransferList");
        arrayTransferGenericMethod = classType.GetMethod("TransferArray");
        dictionaryTransferGenericMethod = classType.GetMethod("TransferDic");
    }

    public static MethodInfo MakeGenericArrayTransferMethod(Type type)
    {
        return arrayTransferGenericMethod.MakeGenericMethod(new Type[] { type });
    }

    public static void AppendIndent(StringBuilder sb, int indent)
    {
        for (int i = 0; i < indent; ++i)
        {
            sb.Append("\t");
        }
    }

    public static bool TransferList<T>(List<T> list, StringBuilder sb, int indent = 0)
    {
        return TransferArray<T>(list.ToArray(), sb, indent);
    }

    public static bool TransferArray<T>(T[] array, StringBuilder sb, int indent = 0)
    {
        int validContentLength = sb.Length;
        AppendIndent(sb, indent);
        sb.Append("{");
        var type = typeof(T);
        bool bSerializeSuc = false;

        if (array.Length <= 0)
        {
            bSerializeSuc = false;
            sb.Remove(sb.Length - validContentLength, validContentLength);
            return bSerializeSuc;
        }

        if (typeof(LuaSerialization).IsAssignableFrom(type))
        {
            int tempValidContentLength = sb.Length;

            ++indent;
            sb.Append("\n");
            Array.ForEach<T>(array, (value) =>
            {
                LuaSerialization serializor = value as LuaSerialization;
                if (serializor.Serialize(sb, indent))
                {
                    bSerializeSuc = true;
                    sb.Append(",\n");
                }
            });
            --indent;

            if (!bSerializeSuc)
                sb.Remove(tempValidContentLength, sb.Length - tempValidContentLength);
            else
                AppendIndent(sb, indent);
        }
        else if (type.IsClass)
        {
            if (type == typeof(string))
            {
                --indent;
                Array.ForEach<T>(array, (value) =>
                {
                    sb.Append("\"");
                    sb.Append(value.ToString().Replace("\n", @"\n").Replace("\"", @"\"""));
                    sb.Append("\"");
                    sb.Append(", ");
                });
                --indent;

                bSerializeSuc = true;
            }
            else
            {
                MethodInfo method = null;
                if (type.GetGenericTypeDefinition() == dictionaryTypeDefinition)
                    method = dictionaryTransferGenericMethod.MakeGenericMethod(type.GetGenericArguments());
                else if (type.GetGenericTypeDefinition() == listTypeDefinition)
                    method = listTransferGenericMethod.MakeGenericMethod(type.GetGenericArguments());
                else if (type.IsArray)
                    method = arrayTransferGenericMethod.MakeGenericMethod(new Type[] { type.GetElementType() });

                if (method != null)
                {
                    int tempValidContentLength = sb.Length;

                    ++indent;
                    sb.Append("\n");
                    Array.ForEach<T>(array, (value) =>
                    {
                        bool bSeirializeResult = (bool)method.Invoke(null, new object[] { value, sb, indent });
                        if (bSeirializeResult)
                        {
                            bSerializeSuc = true;
                            sb.Append(",\n");
                        }
                    });
                    --indent;

                    if (!bSerializeSuc)
                        sb.Remove(tempValidContentLength, sb.Length - tempValidContentLength);
                    else
                        AppendIndent(sb, indent);
                }
            }
        }
        else if (type.IsValueType)
        {
            if (type.IsPrimitive)
            {
                Array.ForEach<T>(array, (value) =>
                {
                    sb.Append(value);
                    sb.Append(", ");
                });

                bSerializeSuc = true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("Can't Serialize Specify Data Type : {0} To Lua", type));
            }
        }

        if (!bSerializeSuc)
        {
            sb.Remove(validContentLength, sb.Length - validContentLength);
        }
        else
        {
            sb.Append("}");
        }

        return bSerializeSuc;
    }

    public static bool TransferDic<T, U>(Dictionary<T, U> dic, StringBuilder sb, int indent = 0)
    {
        int validContentLength = sb.Length;
        AppendIndent(sb, indent);
        sb.Append("{\n");
        var keyType = typeof(T);
        var valueType = typeof(U);
        bool bSerializeSuc = false;

        if (dic.Count <= 0)
        {
            bSerializeSuc = false;
            sb.Remove(sb.Length - validContentLength, validContentLength);
            return bSerializeSuc;
        }

        if (keyType == typeof(float) || keyType == typeof(double) || keyType == typeof(bool))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("Can't Serialize Specify Data Type : {0} To Lua", keyType));
            bSerializeSuc = false;
            sb.Remove(sb.Length - validContentLength, validContentLength);
            return bSerializeSuc;
        }

        bool bStringKey = keyType == typeof(string);
        string strTag = bStringKey ? "\"" : "";
        ++indent;
        AppendIndent(sb, indent);
        foreach (var item in dic)
        {
            int tempValidContentLength = sb.Length;

            sb.Append("[");
            sb.Append(strTag);
            /// 不管是不是自定义数据，只要tostring能用就行
            sb.Append(item.Key);
            sb.Append(strTag);
            sb.Append("] = ");

            if (typeof(LuaSerialization).IsAssignableFrom(valueType))
            {
                sb.Append("\n");
                ++indent;
                LuaSerialization serializor = item.Value as LuaSerialization;
                if (serializor.Serialize(sb, indent))
                {
                    bSerializeSuc = true;
                    sb.Append(",\n");
                }
                --indent;

                if (!bSerializeSuc)
                    sb.Remove(tempValidContentLength, sb.Length - tempValidContentLength);
                else
                    AppendIndent(sb, indent);
            }
            else if (valueType.IsClass)
            {
                if (valueType == typeof(string))
                {
                    sb.Append("\"");
                    sb.Append(item.Value.ToString().Replace("\n", @"\n").Replace("\"", @"\"""));
                    sb.Append("\"");
                    sb.Append(", ");

                    bSerializeSuc = true;
                }
                else
                {
                    MethodInfo method = null;
                    if (valueType.GetGenericTypeDefinition() == dictionaryTypeDefinition)
                        method = dictionaryTransferGenericMethod.MakeGenericMethod(valueType.GetGenericArguments());
                    else if (valueType.GetGenericTypeDefinition() == listTypeDefinition)
                        method = listTransferGenericMethod.MakeGenericMethod(valueType.GetGenericArguments());
                    else if (valueType.IsArray)
                        method = arrayTransferGenericMethod.MakeGenericMethod(new Type[] { valueType.GetElementType() });

                    if (method != null)
                    {
                        sb.Append("\n");
                        ++indent;
                        bool bSeirializeResult = (bool)method.Invoke(null, new object[] { item.Value, sb, indent });
                        if (bSeirializeResult)
                        {
                            bSerializeSuc = true;
                            sb.Append(",\n");
                        }
                        --indent;

                        if (!bSerializeSuc)
                            sb.Remove(tempValidContentLength, sb.Length - tempValidContentLength);
                        else
                            AppendIndent(sb, indent);
                    }
                }
            }
            else if (valueType.IsValueType)
            {
                if (valueType.IsPrimitive)
                {
                    sb.Append(item.Value);
                    sb.Append(", ");

                    bSerializeSuc = true;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("Can't Serialize Specify Data Type : {0} To Lua", valueType));
                }
            }
        }

        if (!bSerializeSuc)
        {
            sb.Remove(validContentLength, sb.Length - validContentLength);
        }
        else
        {
            sb.Append("\n");
            AppendIndent(sb, indent - 1);
            sb.Append("}");
        }

        return bSerializeSuc;
    }
}
