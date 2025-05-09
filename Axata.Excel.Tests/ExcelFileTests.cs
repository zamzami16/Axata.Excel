using Axata.Excel.Domain;
using NUnit.Framework;
using System.Data;
using System.Reflection;

namespace Axata.Excel.Tests;

[TestFixture]
public class ExcelFileTests
{
    private string _validXlsxFile;
    private string _validXlsFile;
    private string _invalidFile;

    [SetUp]
    public void Setup()
    {
        _validXlsxFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString().Replace('_', ' ')}.xlsx");
        _validXlsFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xls");
        _invalidFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.txt");
    }

    [TearDown]
    public void TearDown()
    {
        if (File.Exists(_validXlsxFile)) File.Delete(_validXlsxFile);
        if (File.Exists(_validXlsFile)) File.Delete(_validXlsFile);
        if (File.Exists(_invalidFile)) File.Delete(_invalidFile);
    }

    private static void SetReadonlyProperty(object obj, string propertyName, object value, bool asBackingField = true)
    {
        if (obj != null)
        {
            var type = obj.GetType();
            var fieldName = asBackingField ? $"<{propertyName}>k__BackingField" : propertyName;

            var field = type.GetField(fieldName, BindingFlags.Instance | BindingFlags.NonPublic);
            if (field == null)
                throw new ArgumentException($"Backing field '{fieldName}' not found on '{type.Name}'.");

            field.SetValue(obj, value);
        }
        else
        {
            throw new ArgumentNullException(nameof(obj));
        }
    }

    [Test]
    public void Constructor_Should_Throw_When_FileName_Is_Null_Or_Empty()
    {
        Assert.That(() => new ExcelFile(null),
            Throws.ArgumentNullException
                .With.Property("ParamName").EqualTo("fileName"));

        Assert.That(() => new ExcelFile(""),
            Throws.ArgumentNullException
                .With.Property("ParamName").EqualTo("fileName"));
    }

    [Test]
    public void Constructor_Should_Throw_When_File_Extension_Is_Invalid()
    {
        Assert.That(() => new ExcelFile(_invalidFile),
            Throws.InstanceOf<ArgumentOutOfRangeException>()
                .With.Property("ParamName").EqualTo("fileName")
                .And.Message.Contain("is not recognized as valid"));
    }

    [Test]
    public void Constructor_Should_Not_Throw_For_Valid_Xlsx_File()
    {
        Assert.That(() => new ExcelFile(_validXlsxFile), Throws.Nothing);
    }

    [Test]
    public void Constructor_Should_Not_Throw_For_Valid_Xls_File()
    {
        Assert.That(() => new ExcelFile(_validXlsFile), Throws.Nothing);
    }

    [Test]
    public void ToDataSet_Should_Return_DataSet()
    {
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        var file = ExcelFile.Create(table, _validXlsxFile);
        file.Save();
        var dataSet = file.ToDataSet();

        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(2));
    }

    [Test]
    public void ToDataTable_Should_Return_DataTable()
    {
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        var file = ExcelFile.Create(table, _validXlsxFile);
        file.Save();
        var dataTable = file.ToDataTable();

        Assert.That(dataTable, Is.Not.Null);
        Assert.That(dataTable.Rows.Count, Is.EqualTo(2));
    }

    [Test]
    public void ToDataTable_Should_Throw_When_No_Data_Found()
    {
        var file = new ExcelFile(_validXlsFile);

        Assert.That(() => file.ToDataTable(),
            Throws.TypeOf<FileNotFoundException>()
                .With.Message.Contain("not found"));
    }

    [Test]
    public void Save_Should_Throw_When_DataSource_Not_Set()
    {
        var file = new ExcelFile(_validXlsxFile);

        Assert.That(() => file.Save(),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.EqualTo("Data source is not set."));
    }

    [Test]
    public void SaveAs_Should_Throw_When_FileName_Is_Null_Or_Empty()
    {
        var file = ExcelFile.Create(new DataTable(), _validXlsxFile);

        Assert.That(() => file.SaveAs(null),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.Contain("File name cannot be null or empty."));

        Assert.That(() => file.SaveAs(""),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.Contain("File name cannot be null or empty."));
    }

    [Test]
    public void SaveAs_Should_Throw_When_File_Extension_Is_Invalid()
    {
        var file = ExcelFile.Create(new DataTable(), _validXlsxFile);

        Assert.That(() => file.SaveAs(_invalidFile),
            Throws.InstanceOf<AxataExcelException>()
                .With.Message.Contain("is not recognized as valid"));
    }

    [Test]
    public void SaveAs_Should_Save_File_With_Valid_Extension()
    {
        var file = ExcelFile.Create(new DataTable(), _validXlsxFile);
        var savedPath = file.SaveAs(_validXlsxFile);

        Assert.That(File.Exists(savedPath), Is.True);
    }

    [Test]
    public void Create_Should_Create_File_With_Default_Extension_When_FileName_Not_Provided()
    {
        var tempFile = ExcelFile.Create(new DataTable());
        var fileName = tempFile.Save();

        Assert.That(File.Exists(fileName), Is.True);
        File.Delete(tempFile.FileName);
    }

    [Test()]
    public void ToDataSetTest_ShouldThrows_IfFileNotExists()
    {
        var file = new ExcelFile(_validXlsFile);

        Assert.That(() => file.ToDataSet(),
            Throws.TypeOf<FileNotFoundException>()
                .With.Message.Contain("not found"));
    }

    [Test()]
    [TestCase(".xls")]
    [TestCase(".xlsx")]
    public void ToDataSetTest_ShouldThrows_IfExtension_Invalid(string extension)
    {
        var file = CreateExcelFile(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString().Replace('_', ' ')}{extension}"));

        SetReadonlyProperty(file, "Extension", ".txt");
        Assert.That(() => file.ToDataSet(),
            Throws.TypeOf<ArgumentOutOfRangeException>()
                .With.Message.Contain("is not recognized as valid"));
    }

    private static ExcelFile CreateExcelFile(string excelFile)
    {
        var table = new DataTable();
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");

        var file = ExcelFile.Create(table, excelFile);
        file.Save();
        return file;
    }

    [Test()]
    [TestCaseSource(typeof(ExcelFileTestCase), nameof(ExcelFileTestCase.ExcelFiles))]
    public void ToDataTableTest_FromExcelFile(string file)
    {
        var excelFile = new ExcelFile(file);
        var dataTable = excelFile.ToDataTable();
        Assert.That(dataTable, Is.Not.Null);
        Assert.That(dataTable.Rows, Has.Count.GreaterThanOrEqualTo(2));

        var isValid = ExcelFileTestCase.IsValidAsCustomerData(dataTable);
        Assert.That(isValid, Is.True);
    }

    [Test()]
    [TestCaseSource(typeof(ExcelFileTestCase), nameof(ExcelFileTestCase.ExcelFiles))]
    public void ToDataSetTest_FromExcelFile(string file)
    {
        var excelFile = new ExcelFile(file);
        var ds = excelFile.ToDataSet();

        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables.Count, Is.GreaterThan(0));

        Assert.That(ds.Tables[0].Rows.Count, Is.GreaterThanOrEqualTo(2));
        var isValid = ExcelFileTestCase.IsValidAsCustomerData(ds.Tables[0]);
        Assert.That(isValid, Is.True);
    }

    [Test()]
    [TestCaseSource(typeof(ExcelFileTestCase), nameof(ExcelFileTestCase.ExcelFilesTemp))]
    public void CreateTest_FromDataSet_Success(string fileName)
    {
        var excelFile = ExcelFile.Create(ExcelFileTestCase.CreateDataSet(), fileName);
        var savedFileName = excelFile.Save();

        Assert.That(File.Exists(savedFileName), Is.True);

        var excelFile2 = new ExcelFile(savedFileName);
        var ds = excelFile2.ToDataSet();

        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables.Count, Is.GreaterThan(0));
        Assert.That(ds.Tables[0].Rows.Count, Is.GreaterThanOrEqualTo(2));

        var isValid = ExcelFileTestCase.IsValidAsCustomerData(ds.Tables[0]);
        Assert.That(isValid, Is.True);
    }

    [Test()]
    [TestCaseSource(typeof(ExcelFileTestCase), nameof(ExcelFileTestCase.ExcelFilesTemp))]
    public void CreateTest_FromEnumerableT_Success(string fileName)
    {
        var excelFile = ExcelFile.Create(ExcelFileTestCase.CreateCustomers(), fileName);
        var savedFileName = excelFile.Save();

        Assert.That(File.Exists(savedFileName), Is.True);

        var excelFile2 = new ExcelFile(savedFileName);
        var ds = excelFile2.ToDataSet();

        Assert.That(ds, Is.Not.Null);
        Assert.That(ds.Tables.Count, Is.GreaterThan(0));
        Assert.That(ds.Tables[0].Rows.Count, Is.GreaterThanOrEqualTo(2));

        var isValid = ExcelFileTestCase.IsValidAsCustomerData(ds.Tables[0]);
        Assert.That(isValid, Is.True);
    }

    [Test()]
    public void SaveAsTest_ShouldThrows_WhenError()
    {
        var excelFile = ExcelFile.Create(ExcelFileTestCase.CreateCustomers());
        var excelDataSource = new ExcelDataSource(new DataTable());

        SetReadonlyProperty(excelDataSource, "_dataSource", 1, false);
        SetReadonlyProperty(excelFile, "DataSource", excelDataSource);

        Assert.That(() => excelFile.SaveAs(_validXlsxFile),
            Throws.InstanceOf<AxataExcelException>());
    }
}
