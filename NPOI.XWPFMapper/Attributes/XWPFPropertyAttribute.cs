using System;

namespace NPOI.XWPFMapper.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class XWPFPropertyAttribute : Attribute
    {
        public string XWPFName;
        public XWPFPropertyAttribute(string name) => XWPFName = name;
    }
}
