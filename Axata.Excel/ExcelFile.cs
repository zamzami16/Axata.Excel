using Axata.Excel.Domain;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Axata.Excel;

public sealed class ExcelFile : IExcelFile
{
    static readonly string[] ValidExcelFileExtension = [".xls", ".xlsx",];

    public ExcelFile(string fileName)
    {
        EncodingInitialization.EnsureInitialized();

        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentNullException(nameof(fileName), "File name cannot be null or empty.");

        var extension = Path.GetExtension(fileName);
        if (!ValidExcelFileExtension.Any(ve => ve.Equals(extension, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentOutOfRangeException(nameof(fileName), $"File extension {extension} is not recognized as valid.");
        }

        FileName = fileName;
        Extension = extension;
    }

    public string FileName { get; }

    void EnsureFileExists()
    {
        if (!File.Exists(FileName))
        {
            throw new FileNotFoundException($"File {FileName} not found.");
        }
    }

    public string Extension { get; }

    public DataSet ToDataSet(ExcelDataSetConfiguration config = null)
    {
        EnsureFileExists();

        using FileStream fs = File.Open(FileName, FileMode.Open, FileAccess.Read);
        using IExcelDataReader rdr = Extension switch
        {
            //Old style (2003 or less)
            ".xls" => ExcelReaderFactory.CreateBinaryReader(fs),
            //New style (2007 or higher)
            ".xlsx" => ExcelReaderFactory.CreateOpenXmlReader(fs),
            _ => throw new ArgumentOutOfRangeException(string.Format("File extension {0} is not recognized as valid", Extension)),
        };

        DataSet ds = rdr.AsDataSet(config ?? DefaultDataSetConfiguration);
        return ds;
    }

    public DataTable ToDataTable(ExcelDataSetConfiguration config = null)
    {
        var ds = ToDataSet(config);
        return ds.Tables[0];
    }

    static ExcelDataSetConfiguration DefaultDataSetConfiguration => new()
    {
        UseColumnDataType = true,
        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
        {
            UseHeaderRow = true
        }
    };

    private ExcelDataSource DataSource { get; set; } = default;

    public string Save() => SaveAs(FileName);

    public string SaveAs(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            var err = new ArgumentNullException(nameof(fileName), "File name cannot be null or empty.");
            throw new AxataExcelException(err.Message, err);
        }

        var extension = Path.GetExtension(fileName);

        if (!ValidExcelFileExtension.Any(ve => ve.Equals(extension, StringComparison.OrdinalIgnoreCase)))
        {
            var err = new ArgumentOutOfRangeException(nameof(fileName), $"File extension {extension} is not recognized as valid.");
            throw new AxataExcelException(err.Message, err);
        }

        if (DataSource == default)
        {
            throw new AxataExcelException("Data source is not set.");
        }

        try
        {
            File.WriteAllBytes(FileName, DataSource.ReadExcelByte());
            return fileName;
        }
        catch (Exception err)
        {
            throw new AxataExcelException(err.Message, err);
        }
    }

    public static ExcelFile Create<T>(IEnumerable<T> items, string fileName = "")
        where T : class
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        }

        return new ExcelFile(fileName)
        {
            DataSource = new ExcelDataSource<T>(items),
        };
    }

    public static ExcelFile Create(DataSet dataSet, string fileName = "")
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        }

        return new ExcelFile(fileName)
        {
            DataSource = new ExcelDataSource(dataSet),
        };
    }

    public static ExcelFile Create(DataTable dataTable, string fileName = "")
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        }

        return new ExcelFile(fileName)
        {
            DataSource = new ExcelDataSource(dataTable),
        };
    }
}
