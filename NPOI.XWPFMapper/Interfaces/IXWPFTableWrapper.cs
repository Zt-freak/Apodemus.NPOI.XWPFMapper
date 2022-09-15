using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Attributes;
using NPOI.XWPFMapper.Managers;
using System.Linq;
using System.Reflection;
using System;

namespace NPOI.XWPFMapper.Interfaces
{
    public interface IXWPFTableWrapper<T> where T : IXWPFMappable
    {
        Type ObjectType { get; }
        XWPFTable Table { get; }
        void AddRow(IXWPFMappable mappableObject);
    }
}
