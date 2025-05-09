using Axata.Excel.Tests.Resources;
using NUnit.Framework;
using System.Data;
using System.Reflection;

namespace Axata.Excel.Tests;

internal static class ExcelFileTestCase
{
    internal static IEnumerable<string> ExcelFiles()
    {
        var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";

        yield return Path.Combine(basePath, "Resources", Customers.CustomerXls);
        yield return Path.Combine(basePath, "Resources", Customers.CustomerXlsx);
    }

    internal static IEnumerable<string> ExcelFilesTemp()
    {
        var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";

        yield return "";
        yield return Path.Combine(basePath, Customers.CustomerXlsx);
    }

    static readonly string[] CustomerColumnName = [
        "Name",
        "Age",
    ];

    internal static bool IsValidAsCustomerData(DataTable table)
    {
        var validColumn = table.Columns.Cast<DataColumn>()
            .Select(c => c.ColumnName)
            .All(c => CustomerColumnName.Contains(c));

        List<(string, double)> expected =
        [
            ("Alice", 15),
            ("Bob", 60),
        ];

        var actual = table.AsEnumerable()
            .Select(r => (r.Field<string>("Name"), r.Field<double>("Age")))
            .ToList();

        try
        {
            Assert.That(actual, Is.EquivalentTo(expected));
            Assert.That(validColumn, Is.True);
        }
        catch (Exception)
        {
            return false;
        }

        return true;
    }

    public static DataSet CreateDataSet()
    {
        var ds = new DataSet();
        ds.Tables.Add(CreateDataTable());
        return ds;
    }

    public static DataTable CreateDataTable()
    {
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Columns.Add("Age", typeof(double));

        table.Rows.Add("Alice", 15);
        table.Rows.Add("Bob", 60);

        return table;
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public double Age { get; set; }
    }

    public static IEnumerable<Customer> CreateCustomers()
    {
        return
        [
            new() { Name = "Alice", Age = 15 },
            new() { Name = "Bob", Age = 60 },
        ];
    }
}
