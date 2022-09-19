using ConsoleApp1;
using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Enums;
using NPOI.XWPFMapper.Managers;

XWPFDocument document = new XWPFDocument();

ExampleClass exampleData = new ExampleClass()
{
    Enum = ExampleEnum.Red,
    Name = "This is a test",
    IgnoredMember = "This member will be ignored",
    Address = new ExampleChildClass()
    {
        Address = "1",
        City = "1",
        CountryCode = "1",
        XWPFTableAlignment = XWPFTableAlignment.Column
    },
    Address2 = new ExampleChildClass()
    {
        Address = "2",
        City = "2",
        CountryCode = "2"
    },
    Address3 = new ExampleChildClass()
    {
        Address = "3",
        City = "3",
        CountryCode = "3"
    }
};

// XWPFTableWrapper requires a Type argument and an XWPFDocument object to work
// In addition you can optionally set the direction of the table with an enum XWPFTableAlignment (default is Row)

XWPFTableWrapper<ExampleClass> wrapper = new XWPFTableWrapper<ExampleClass>(document, XWPFTableAlignment.Row);
wrapper.Insert(exampleData);

using FileStream fs = new FileStream("test.docx", FileMode.Create);
document.Write(fs);