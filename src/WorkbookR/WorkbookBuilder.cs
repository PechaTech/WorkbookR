using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WorkbookR;

public static class WorkbookBuilder
{
    public static WorkbookConfigurator<T> ForModel<T>()
    {
        return new WorkbookConfigurator<T>();
    }
}

public class WorkbookConfigurator<T>
{
    public XLWorkbook Build()
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(typeof(T).Name);
        
        foreach (var property in typeof(T).GetProperties().Select((value, i) => ( value, i  )))
        {
            worksheet.Cell(1, property.i + 1).Value = property.value.Name;
        }

        return workbook;
    }
}