// See https://aka.ms/new-console-template for more information

using WorkbookR;

var models = new TestModel[]
{
    new() { FooString = "test1", BarInt = 1 },
    new() { FooString = "test2", BarInt = 2 }
};
var book = WorkbookBuilder.For(models).Build();
book.SaveAs("../../../HelloWorld.xlsx");
book.Dispose();

public class TestModel
{
    public string FooString { get; set; }
    public int BarInt { get; set; }
}