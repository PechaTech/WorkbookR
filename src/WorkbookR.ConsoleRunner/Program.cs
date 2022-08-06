// See https://aka.ms/new-console-template for more information

using WorkbookR;

var book = WorkbookBuilder.ForModel<TestModel>().Build();
book.SaveAs("../../../HelloWorld.xlsx");
book.Dispose();

public class TestModel
{
    public string FooString { get; set; }
    public int BarInt { get; set; }
}