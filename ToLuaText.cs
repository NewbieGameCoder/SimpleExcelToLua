using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

public interface LuaSerialization
{
    void Serialize(StringBuilder sb, int indent);
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

    public static void TransferList<T>(List<T> list, StringBuilder sb, int indent = 0)
    {
        TransferArray<T>(list.ToArray(), sb, indent);
    }

    public static void TransferArray<T>(T[] array, StringBuilder sb, int indent = 0)
    {
        AppendIndent(sb, indent);
        sb.Append("{");
        var type = typeof(T);

        if (typeof(LuaSerialization).IsAssignableFrom(type))
        {
            ++indent;
            sb.Append("\n");
            Array.ForEach<T>(array, (value) =>
            {
                LuaSerialization serializor = value as LuaSerialization;
                serializor.Serialize(sb, indent);
                sb.Append(",\n");
            });
            --indent;
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
                    ++indent;
                    sb.Append("\n");
                    Array.ForEach<T>(array, (value) =>
                    {
                        method.Invoke(null, new object[] { value, sb, indent });
                        sb.Append(",\n");
                    });
                    --indent;
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
            }
            else
            {
                Console.WriteLine(string.Format("Can't Serialize Specify Data Type : {0} To Lua", type));
            }
        }

        sb.Append("}");
    }

    public static void TransferDic<T, U>(Dictionary<T, U> dic, StringBuilder sb, int indent = 0)
    {
        AppendIndent(sb, indent);
        sb.Append("{\n");
        var keyType = typeof(T);
        var valueType = typeof(U);

        bool bStringKey = keyType == typeof(string);
        string strTag = bStringKey ? "\"" : "";
        ++indent;
        AppendIndent(sb, indent);
        foreach (var item in dic)
        {
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
                serializor.Serialize(sb, indent);
                --indent;
                sb.Append(",\n");
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
                        method.Invoke(null, new object[] { item.Value, sb, indent });
                        --indent;
                        sb.Append(",\n");
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
                }
                else
                {
                    Console.WriteLine(string.Format("Can't Serialize Specify Data Type : {0} To Lua", valueType));
                }
            }
        }

        sb.Append("\n");
        AppendIndent(sb, indent - 1);
        sb.Append("}");
    }
}
