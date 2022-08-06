using ClosedXML.Excel;

namespace WorkbookR;

public static class WorkbookBuilder
{
    public static WorkbookConfigurator<T> For<T>()
    {
        return new WorkbookConfigurator<T>();
    }
    
    public static WorkbookConfigurator<T> For<T>(IEnumerable<T> data)
    {
        return new WorkbookConfigurator<T>(data);
    }
}

public class WorkbookConfigurator<T>
{
    public IEnumerable<T> Data { get; }
    
    public WorkbookConfigurator()
    {
        Data = Enumerable.Empty<T>();
    }

    public WorkbookConfigurator(IEnumerable<T> data)
    {
        Data = data;
    }
    
    public XLWorkbook Build()
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(typeof(T).Name);
        
        foreach (var property in typeof(T).GetProperties().Select((value, i) => (value, i)))
        {
            worksheet.Cell(1, property.i + 1).Value = property.value.Name;
        }

        foreach (var item in Data.Select((value, i) => (value, i)))
        {
            foreach (var property in typeof(T).GetProperties().Select((value, i) => (value, i)))
            {
                worksheet.Cell(2 + item.i, property.i + 1).Value =
                    Convert.ChangeType(item.value?.GetType().GetProperty(property.value.Name)?.GetValue(item.value), property.value.PropertyType);
            }
        }

        return workbook;
    }
}