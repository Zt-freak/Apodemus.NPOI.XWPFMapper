using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Attributes;
using NPOI.XWPFMapper.Interfaces;
using System;
using System.Linq;
using System.Reflection;

namespace NPOI.XWPFMapper.Managers
{
    public class XWPFTableWrapper<T> : IXWPFTableWrapper<T> where T : IXWPFMappable
    {
        public Type ObjectType { get; } = typeof(T);
        private readonly PropertyInfo[] _objectProperties = typeof(T).GetProperties();
        public XWPFTable Table { get; }

        public XWPFTableWrapper(XWPFDocument document)
        {
            try
            {
                Table = document.CreateTable();

                Map();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public XWPFTableWrapper(XWPFTableCell tableCell)
        {
            try
            {
                CT_Tbl ctTbl = tableCell.GetCTTc().AddNewTbl();
                tableCell.GetCTTc().AddNewP();

                Table = new XWPFTable(ctTbl, tableCell);

                Map();
                tableCell.RemoveParagraph(0);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void Map(PropertyInfo[] properties = null)
        {
            try
            {
                if (properties is null)
                    properties = _objectProperties;

                bool firstCycle = true;
                foreach (PropertyInfo propertyInfo in properties.Where(p => p.CustomAttributes.Any(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)))))
                {
                    XWPFPropertyAttribute attr = (XWPFPropertyAttribute)Attribute.GetCustomAttribute(propertyInfo, typeof(XWPFPropertyAttribute));
                    if (firstCycle)
                        firstCycle = false;
                    else
                    {
                        XWPFTableRow newRow = new XWPFTableRow(new CT_Row(), Table);
                        newRow.AddNewTableCell();
                        Table.AddRow(newRow);
                    }
                    Table.GetRow(Table.Rows.Count - 1).GetCell(0).SetText(attr.XWPFName);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void AddRow(IXWPFMappable mappableObject)
        {
            try
            {
                if (mappableObject != null && mappableObject.GetType() != typeof(T))
                    throw new ArgumentException("mappableObject types do not match");

                InsertRowValue(mappableObject);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void InsertRowValue(IXWPFMappable mappableObject)
        {
            try
            {
                int index = 0;
                foreach (PropertyInfo propertyInfo in _objectProperties.Where(p => p.CustomAttributes.Any(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)))))
                {
                    CustomAttributeData attr = propertyInfo.CustomAttributes.First(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)));
                    if (typeof(IXWPFMappable).IsAssignableFrom(propertyInfo.PropertyType))
                    {
                        XWPFTableCell newCell = Table.GetRow(index).CreateCell();

                        var value = (IXWPFMappable)Convert.ChangeType(propertyInfo.GetValue(mappableObject), propertyInfo.PropertyType);

                        if (value == null)
                            continue;

                        Type type = typeof(XWPFTableWrapper<>).MakeGenericType(propertyInfo.PropertyType);
                        dynamic newTableWrapper = Activator.CreateInstance(type, new object[] { newCell });

                        MethodInfo addRowMethod = ((object)newTableWrapper).GetType().GetMethod("AddRow");
                        addRowMethod.Invoke(newTableWrapper, new object[] { value });
                    }
                    else
                    {
                        Table.GetRow(index).CreateCell().SetText(propertyInfo.GetValue(mappableObject)?.ToString());
                    }
                    index++;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
