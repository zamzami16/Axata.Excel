# Axata.Excel

[![NuGet](https://img.shields.io/nuget/v/Axata.Excel.svg)](https://www.nuget.org/packages/Axata.Excel/)
[![Build Status](https://github.com/zamzami16/Axata.Excel/actions/workflows/test.yml/badge.svg)](https://github.com/zamzami16/Axata.Excel/actions)

Axata.Excel is a .NET library designed for easy and efficient Excel file manipulation. It provides a simple yet powerful API for reading, writing, and transforming Excel files.

## Installation

To install the library via NuGet, run the following command:

```bash
dotnet add package Axata.Excel
````

Or add the following line to your `.csproj` file:

```xml
<PackageReference Include="Axata.Excel" Version="0.0.1" />
```

## Features

* Read and write Excel files (`.xlsx`, `.xls`)
* Convert Excel worksheets to DataSets and DataTables
* Flexible configuration for data extraction

## Getting Started

### Basic Example

Here is a simple example of how to use `Axata.Excel` to read an Excel file into a `DataTable`:

```csharp
using Axata.Excel;
using System.Data;

// Load an Excel file
IExcelFile excelFile = new ExcelFile("example.xlsx");

// Convert the first sheet to a DataTable
DataTable dataTable = excelFile.ToDataTable();

// Print data from the DataTable
foreach (DataRow row in dataTable.Rows)
{
    Console.WriteLine(string.Join(", ", row.ItemArray));
}
```

### Saving an Excel File

```csharp
using Axata.Excel;

// Create a new Excel file
IExcelFile excelFile = new ExcelFile("my-report.xlsx");

// Save the file
excelFile.Save();
```

### Convert to DataSet

```csharp
using Axata.Excel;
using System.Data;

// Load the file and convert it to a DataSet
IExcelFile excelFile = new ExcelFile("complex-report.xlsx");
DataSet dataSet = excelFile.ToDataSet();

// Access individual tables (worksheets)
DataTable firstSheet = dataSet.Tables[0];
```

## Configuration

You can customize the way Excel files are read by providing an `ExcelDataSetConfiguration`:

```csharp
var config = new ExcelDataSetConfiguration
{
    UseHeaderRow = true
};

DataSet dataSet = excelFile.ToDataSet(config);
```

## API Reference

### IExcelFile Interface

* **Extension** - Gets the file extension.
* **FileName** - Gets the file name without the directory path.
* **Save()** - Saves the Excel file to its current location.
* **SaveAs(string fileName)** - Saves the Excel file with a specified name.
* **ToDataSet(ExcelDataSetConfiguration config = null)** - Converts the Excel file to a `DataSet`.
* **ToDataTable(ExcelDataSetConfiguration config = null)** - Converts the first worksheet to a `DataTable`.

## Building from Source

Clone the repository:

```bash
git clone https://github.com/zamzami16/Axata.Excel.git
cd Axata.Excel
dotnet build
```
