using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using NPOI.XWPFMapper.Attributes;
using NPOI.XWPFMapper.Enums;
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
        public XWPFTableAlignment XWPFTableAlignment { get; }
        public XWPFTable Table { get; }

        public XWPFTableWrapper(XWPFDocument document, XWPFTableAlignment alignment = XWPFTableAlignment.Row)
        {
            XWPFTableAlignment = alignment;
            Table = document.CreateTable();

            Map();
        }

        public XWPFTableWrapper(XWPFTableCell tableCell, XWPFTableAlignment alignment = XWPFTableAlignment.Row)
        {
            XWPFTableAlignment = alignment;
            CT_Tbl ctTbl = tableCell.GetCTTc().AddNewTbl();
            tableCell.GetCTTc().AddNewP();

            Table = new XWPFTable(ctTbl, tableCell);

            Map();
            tableCell.RemoveParagraph(0);
        }

        private void Map(PropertyInfo[] properties = null)
        {
            if (properties is null)
                properties = _objectProperties;

            switch (XWPFTableAlignment)
            {
                case XWPFTableAlignment.Column:
                    MapColumns(properties);
                    break;
                case XWPFTableAlignment.Row:
                default:
                    MapRows(properties);
                    break;
            }
        }

        private void MapColumns(PropertyInfo[] properties)
        {
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

        private void MapRows(PropertyInfo[] properties)
        {
            bool firstCycle = true;
            foreach (PropertyInfo propertyInfo in properties.Where(p => p.CustomAttributes.Any(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)))))
            {
                XWPFPropertyAttribute attr = (XWPFPropertyAttribute)Attribute.GetCustomAttribute(propertyInfo, typeof(XWPFPropertyAttribute));
                if (firstCycle)
                    firstCycle = false;
                else
                {
                    Table.Rows[0].CreateCell();
                }
                Table.Rows[0].GetTableCells().Last().SetText(attr.XWPFName);
            }
        }

        public void Insert(IXWPFMappable mappableObject)
        {
            if (mappableObject != null && mappableObject.GetType() != typeof(T))
                throw new ArgumentException("mappableObject types do not match");

            switch (XWPFTableAlignment)
            {
                case XWPFTableAlignment.Column:
                    InsertColumnValue(mappableObject);
                    break;
                case XWPFTableAlignment.Row:
                default:
                    InsertRowValue(mappableObject);
                    break;
            }
        }

        private void InsertColumnValue(IXWPFMappable mappableObject)
        {
            try
            {
                int index = 0;
                foreach (PropertyInfo propertyInfo in _objectProperties.Where(p => p.CustomAttributes.Any(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)))))
                {
                    XWPFPropertyAttribute attr = (XWPFPropertyAttribute)Attribute.GetCustomAttribute(propertyInfo, typeof(XWPFPropertyAttribute));
                    if (typeof(IXWPFMappable).IsAssignableFrom(propertyInfo.PropertyType))
                    {
                        XWPFTableCell newCell = Table.GetRow(index).CreateCell();

                        var value = (IXWPFMappable)Convert.ChangeType(propertyInfo.GetValue(mappableObject), propertyInfo.PropertyType);

                        if (value == null)
                            continue;

                        XWPFTableAlignment alignment;
                        if (value.XWPFTableAlignment != XWPFTableAlignment.NotSet)
                            alignment = XWPFTableAlignment.Column;
                        else if (attr.TableAlignment != XWPFTableAlignment.NotSet)
                            alignment = attr.TableAlignment;
                        else if (XWPFTableAlignment != XWPFTableAlignment.NotSet)
                            alignment = XWPFTableAlignment;
                        else
                            alignment = XWPFTableAlignment.Column;

                        Type type = typeof(XWPFTableWrapper<>).MakeGenericType(propertyInfo.PropertyType);
                        dynamic newTableWrapper = Activator.CreateInstance(type, new object[] { newCell, alignment });

                        MethodInfo addRowMethod = ((object)newTableWrapper).GetType().GetMethod("Insert");
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

        private void InsertRowValue(IXWPFMappable mappableObject)
        {
            try
            {
                int index = 0;
                Table.AddRow(new XWPFTableRow(new CT_Row(), Table));
                foreach (PropertyInfo propertyInfo in _objectProperties.Where(p => p.CustomAttributes.Any(a => a.AttributeType.Equals(typeof(XWPFPropertyAttribute)))))
                {
                    XWPFPropertyAttribute attr = (XWPFPropertyAttribute)Attribute.GetCustomAttribute(propertyInfo, typeof(XWPFPropertyAttribute));
                    if (typeof(IXWPFMappable).IsAssignableFrom(propertyInfo.PropertyType))
                    {
                        XWPFTableCell newCell = Table.Rows.Last().CreateCell();

                        var value = (IXWPFMappable)Convert.ChangeType(propertyInfo.GetValue(mappableObject), propertyInfo.PropertyType);

                        if (value == null)
                            continue;

                        XWPFTableAlignment alignment;
                        if (value.XWPFTableAlignment != XWPFTableAlignment.NotSet)
                            alignment = XWPFTableAlignment.Column;
                        else if (attr.TableAlignment != XWPFTableAlignment.NotSet)
                            alignment = attr.TableAlignment;
                        else if (XWPFTableAlignment != XWPFTableAlignment.NotSet)
                            alignment = XWPFTableAlignment;
                        else
                            alignment = XWPFTableAlignment.Row;

                        Type type = typeof(XWPFTableWrapper<>).MakeGenericType(propertyInfo.PropertyType);
                        dynamic newTableWrapper = Activator.CreateInstance(type, new object[] { newCell, alignment });

                        MethodInfo addColumnMethod = ((object)newTableWrapper).GetType().GetMethod("Insert");
                        addColumnMethod.Invoke(newTableWrapper, new object[] { value });
                    }
                    else
                    {
                        Table.Rows.Last().CreateCell().SetText(propertyInfo.GetValue(mappableObject)?.ToString());
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
