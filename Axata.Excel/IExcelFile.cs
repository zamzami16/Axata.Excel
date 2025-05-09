using ExcelDataReader;
using System.Data;

namespace Axata.Excel;

/// <summary>
/// Represents a general interface for working with Excel files,
/// providing methods to save, export, and convert Excel data.
/// </summary>
public interface IExcelFile
{
    /// <summary>
    /// Gets the file extension of the Excel file (e.g., ".xlsx" or ".xls").
    /// </summary>
    string Extension { get; }

    /// <summary>
    /// Gets the name of the Excel file without the directory path.
    /// </summary>
    string FileName { get; }

    /// <summary>
    /// Saves the Excel file to its current location.
    /// </summary>
    /// <returns>
    /// The full path of the saved Excel file.
    /// </returns>
    string Save();

    /// <summary>
    /// Saves the Excel file with a specified file name.
    /// </summary>
    /// <param name="fileName">The full path or file name to save the file as.</param>
    /// <returns>
    /// The full path of the saved Excel file.
    /// </returns>
    string SaveAs(string fileName);

    /// <summary>
    /// Converts the Excel file to a <see cref="DataSet"/> representation.
    /// </summary>
    /// <param name="config">
    /// Optional configuration for controlling how the data is extracted from the Excel file.
    /// If not provided, default settings are used.
    /// </param>
    /// <returns>
    /// A <see cref="DataSet"/> containing the data from the Excel file, 
    /// where each worksheet is represented as a <see cref="DataTable"/>.
    /// </returns>
    DataSet ToDataSet(ExcelDataSetConfiguration config = null);

    /// <summary>
    /// Converts the Excel file to a single <see cref="DataTable"/> representation.
    /// </summary>
    /// <param name="config">
    /// Optional configuration for controlling how the data is extracted from the Excel file.
    /// If not provided, default settings are used.
    /// </param>
    /// <returns>
    /// A <see cref="DataTable"/> containing the data from the first worksheet of the Excel file.
    /// </returns>
    DataTable ToDataTable(ExcelDataSetConfiguration config = null);
}
