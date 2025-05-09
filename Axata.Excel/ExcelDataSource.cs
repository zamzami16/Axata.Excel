using Axata.Excel.Domain;
using Axata.Excel.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Axata.Excel;

public class ExcelDataSource
{
    private readonly object _dataSource;

    public ExcelDataSource(object dataSource)
    {
        EncodingInitialization.EnsureInitialized();

        _dataSource = dataSource
            ?? throw new ArgumentNullException(nameof(dataSource));

        if (_dataSource is not DataTable
             && _dataSource is not DataSet
             && _dataSource is not IEnumerable)
        {
            var err = new ArgumentException(
                "Unsupported data source type. Data source must be a DataTable, DataSet, or IEnumerable.",
                nameof(dataSource));
            throw new AxataExcelException(err.Message, err);
        }
    }

    public virtual byte[] ReadExcelByte()
    {
        return _dataSource switch
        {
            DataTable dt => dt.ToExcel(),
            DataSet ds => ds.ToExcel(),
            IEnumerable e =>
                e.Cast<object>().ToExcel(),
            _ => throw new AxataExcelException(
                "Unsupported data source type. Data source must be a DataTable, DataSet, or IEnumerable.")
        };
    }
}

public class ExcelDataSource<T>(IEnumerable<T> dataSource) : ExcelDataSource(dataSource) where T : class
{
    public override byte[] ReadExcelByte()
    {
        return dataSource.ToExcel();
    }

    public static ExcelDataSource<T> Create(IEnumerable<T> dataSource)
    {
        return new ExcelDataSource<T>(dataSource);
    }
}
