using Axata.Excel.Domain;
using NUnit.Framework;
using System.Data;
using System.Reflection;

namespace Axata.Excel.Tests;

[TestFixture]
public class ExcelDataSourceTests
{
    [Test]
    public void Constructor_Should_Throw_When_DataSource_Is_Null()
    {
        Assert.That(() => new ExcelDataSource(null),
            Throws.ArgumentNullException
                .With.Property("ParamName").EqualTo("dataSource"));
    }

    [Test]
    public void Constructor_Should_Throw_When_DataSource_Is_Invalid_Type()
    {
        Assert.That(() => new ExcelDataSource(123),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.Contain("Data source must be a DataTable, DataSet, or IEnumerable"));
    }

    [Test]
    public void Constructor_Should_Not_Throw_For_DataTable()
    {
        var table = new DataTable();
        Assert.That(() => new ExcelDataSource(table), Throws.Nothing);
    }

    [Test]
    public void Constructor_Should_Not_Throw_For_DataSet()
    {
        var dataSet = new DataSet();
        Assert.That(() => new ExcelDataSource(dataSet), Throws.Nothing);
    }

    [Test]
    public void Constructor_Should_Not_Throw_For_IEnumerable()
    {
        var list = new List<string> { "A", "B", "C" };
        Assert.That(() => new ExcelDataSource(list), Throws.Nothing);
    }

    [Test]
    public void ReadExcelByte_Should_Throw_For_Unsupported_DataSource_Type()
    {
        // Arrange
        var excelDataSource = new ExcelDataSource(new DataTable());

        // Use reflection to set the private _dataSource field to an unsupported type
        var field = typeof(ExcelDataSource).GetField("_dataSource", BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.That(field, Is.Not.Null);
        field.SetValue(excelDataSource, new MemoryStream());

        // Act & Assert
        Assert.That(() => excelDataSource.ReadExcelByte(),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.Contain("Unsupported data source type"));
    }

    [Test]
    public void ReadExcelByte_Should_Work_For_DataTable()
    {
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        var source = new ExcelDataSource(table);
        var bytes = source.ReadExcelByte();

        Assert.That(bytes, Is.Not.Empty);
    }

    [Test]
    public void ReadExcelByte_Should_Work_For_DataSet()
    {
        var dataSet = new DataSet();
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");
        dataSet.Tables.Add(table);

        var source = new ExcelDataSource(dataSet);
        var bytes = source.ReadExcelByte();

        Assert.That(bytes, Is.Not.Empty);
    }

    public sealed class ExcelItems
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Name { get; set; } = string.Empty;
    }

    [Test]
    public void ReadExcelByte_Should_Work_For_IEnumerable_Of_Custom_Type()
    {
        var list = new List<ExcelItems>
        {
            new() { Name = "Alice" },
            new() { Name = "Bob" }
        };

        var source = new ExcelDataSource(list);
        var bytes = source.ReadExcelByte();
        Assert.That(bytes, Is.Not.Empty);
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; } = 0;
    }

    [Test]
    public void GenericConstructor_ShouldAcceptClassCollection()
    {
        // Arrange
        var customers = new List<Customer>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 40 }
        };

        // Act
        var dataSource = new ExcelDataSource<Customer>(customers);
        byte[] bytes = dataSource.ReadExcelByte();

        // Assert
        Assert.That(bytes, Is.Not.Null);
        Assert.That(bytes, Is.Not.Empty);
    }

    [Test]
    public void GenericConstructor_ShouldAcceptEmptyClassCollection()
    {
        // Arrange
        var emptyList = new List<Customer>();

        // Act
        var dataSource = new ExcelDataSource<Customer>(emptyList);
        byte[] bytes = dataSource.ReadExcelByte();

        // Assert
        Assert.That(bytes, Is.Not.Null);
        Assert.That(bytes, Is.Not.Empty);
    }

    [Test]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Major Code Smell", "S4144:Methods should not have identical implementations", Justification = "<Pending>")]
    public void GenericReadExcelByte_ShouldReturnNonEmptyByteArrayForValidData()
    {
        // Arrange
        var customers = new List<Customer>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 40 }
        };

        var dataSource = new ExcelDataSource<Customer>(customers);

        // Act
        byte[] bytes = dataSource.ReadExcelByte();

        // Assert
        Assert.That(bytes, Is.Not.Null);
        Assert.That(bytes, Is.Not.Empty);
    }

    [Test]
    public void GenericCreate_ShouldReturnInstanceOfExcelDataSource()
    {
        // Arrange
        var customers = new List<Customer>();

        // Act
        var dataSource = ExcelDataSource<Customer>.Create(customers);

        // Assert
        Assert.That(dataSource, Is.Not.Null);
        Assert.That(dataSource, Is.InstanceOf<ExcelDataSource<Customer>>());
    }

    [Test]
    public void ReadExcelByte_Generic_OverNonGeneric_ShouldInvoked_FromGeneric()
    {
        var customers = new List<Customer>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 40 }
        };

        var dataSource = new ExcelDataSource<Customer>(customers);
        ExcelDataSource source = dataSource;
        byte[] bytes = source.ReadExcelByte();
        Assert.That(bytes, Is.Not.Null);
        Assert.That(bytes, Is.Not.Empty);
    }
}
