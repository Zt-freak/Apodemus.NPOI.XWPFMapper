using NPOI.XWPFMapper.Attributes;
using NPOI.XWPFMapper.Interfaces;

namespace NPOI.XWPFMapper.ConsoleApp
{
    internal class ExampleClass : IXWPFMappable
    {
        [XWPFProperty("Color")]
        public ExampleEnum Enum { get; set; }
        [XWPFProperty("Name")]
        public string Name { get; set; }
        [XWPFProperty("Address")]
        public ExampleChildClass Address { get; set; }
        public string IgnoredMember { get; set; }
    }

    internal enum ExampleEnum
    {
        Blue,
        Red,
        Green
    }

    internal class ExampleChildClass : IXWPFMappable
    {
        [XWPFProperty("Street")]
        public string Address { get; set; }
        [XWPFProperty("Place")]
        public string City { get; set; }
        [XWPFProperty("Country")]
        public string CountryCode { get; set; }
    }
}
