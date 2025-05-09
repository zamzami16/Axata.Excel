using System.Reflection;

namespace Axata.Excel.Utils;

public sealed class ColumnInfo
{
    internal ColumnInfo(int index, ColumnSchema schema)
    {
        Index = index;
        Schema = schema;
    }

    internal ColumnSchema Schema { get; }

    public int Index { get; }
    public string Name { get => Schema.Name; }
    public MemberInfo Member { get => Schema.Member; }
}
