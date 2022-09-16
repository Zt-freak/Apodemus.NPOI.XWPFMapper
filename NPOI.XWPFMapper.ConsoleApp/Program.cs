using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Enums;
using NPOI.XWPFMapper.Managers;
using System.IO;

namespace NPOI.XWPFMapper.ConsoleApp
{
    class Program
    {
        static void Main()
        {
            XWPFDocument document = new XWPFDocument();

            ExampleClass exampleData = new ExampleClass()
            {
                Enum = ExampleEnum.Red,
                Name = "This is a test",
                IgnoredMember = "This member will be ignored",
                Address = new ExampleChildClass()
                {
                    Address = "Burgemeester Schönfeldplein",
                    City = "Winschoten",
                    CountryCode = "NL"
                }
            };

            XWPFTableWrapper<ExampleClass> rowTableWrapper = new XWPFTableWrapper<ExampleClass>(document);
            rowTableWrapper.Insert(exampleData);

            document.CreateParagraph();

            XWPFTableWrapper<ExampleClass> columnTableWrapper = new XWPFTableWrapper<ExampleClass>(document, XWPFTableAlignment.Column);
            columnTableWrapper.Insert(exampleData);

            using FileStream fs = new FileStream("test.docx", FileMode.Create);
            document.Write(fs);
        }
    }
}
