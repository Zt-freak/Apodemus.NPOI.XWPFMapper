using NPOI.XWPFMapper.Enums;
using System;

namespace NPOI.XWPFMapper.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class XWPFPropertyAttribute : Attribute
    {
        public string XWPFName;
        public XWPFTableAlignment TableAlignment;
        public XWPFPropertyAttribute(string name) => XWPFName = name;
        public XWPFPropertyAttribute(string name, XWPFTableAlignment alignment)
        {
            XWPFName = name;
            TableAlignment = alignment;
        }
    }
}
