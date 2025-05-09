namespace Axata.Excel.Domain;

[ExcludeFromCodeCoverage]
public class AxataExcelException : Exception
{
    public AxataExcelException(string message) : base(message)
    {
    }

    public AxataExcelException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
