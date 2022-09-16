using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Enums;
using System;

namespace NPOI.XWPFMapper.Interfaces
{
    public interface IXWPFTableWrapper<T> where T : IXWPFMappable
    {
        Type ObjectType { get; }
        XWPFTableAlignment XWPFTableAlignment  { get; }
        XWPFTable Table { get; }
        void Insert(IXWPFMappable mappableObject);
    }
}
