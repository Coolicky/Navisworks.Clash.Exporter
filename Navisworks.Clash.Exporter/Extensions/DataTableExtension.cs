using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using Navisworks.Clash.Exporter.Data;
using Navisworks.Clash.Exporter.Extensions.Attributes;

namespace Navisworks.Clash.Exporter.Extensions
{
    public static class DataTableExtension
    {
        public static DataTable ToDataTable<T>(this IEnumerable<T> list)
        {
            var dt = new DataTable();
            var enumerable = list as T[] ?? list.ToArray();

            if (typeof(T).GetCustomAttributes(typeof(TableNameAttribute), false).FirstOrDefault() is TableNameAttribute
                attribute)
            {
                dt.TableName = attribute.Name;
            }
            else
            {
                dt.TableName = typeof(T).Name;
            }

            var allProperties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var properties = allProperties
                .Where(r =>
                    !(r.GetCustomAttributes(typeof(IgnoreColumnAttribute), false).FirstOrDefault() is
                        IgnoreColumnAttribute))
                .ToList();
            foreach (var propertyInfo in properties)
            {
                var name = propertyInfo.Name;
                if (propertyInfo.GetCustomAttributes(typeof(ColumnNameAttribute), false).FirstOrDefault() is
                    ColumnNameAttribute
                    columnNameAttribute)
                {
                    name = columnNameAttribute.Name;
                }

                dt.Columns.Add(name, Nullable.GetUnderlyingType(
                    propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
            }

            foreach (var element in enumerable)
            {
                var props = new List<object>();
                props.AddRange(properties.Select(propertyInfo => propertyInfo.GetValue(element)));
                dt.Rows.Add(props.ToArray());
            }

            return dt;
        }

        public static DataTable ToElementDataTable(this List<ElementDto> elements)
        {
            var dt = new DataTable();

            dt.TableName = "Elements";

            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Guid", typeof(string));
            dt.Columns.Add("ClassName", typeof(string));
            dt.Columns.Add("Model", typeof(string));

            var dictionaries = elements.Select(r => r.QuickProperties).ToList();
            foreach (var dictionary in dictionaries)
            {
                foreach (var keyValuePair in dictionary)
                {
                    if (!dt.Columns.Contains(keyValuePair.Key))
                    {
                        dt.Columns.Add(keyValuePair.Key);
                    }
                }
            }

            foreach (var element in elements
                         .OrderByDescending(r => r.QuickProperties.Count)
                         .GroupBy(r => r.Guid)
                         .Select(g => g.First()))
            {
                var row = dt.NewRow();
                row["Name"] = element.Name;
                row["Guid"] = element.Guid;
                row["ClassName"] = element.ClassName;
                row["Model"] = element.SourceFile;
                foreach (var keyValuePair in element.QuickProperties)
                {
                    row[keyValuePair.Key] = keyValuePair.Value;
                }

                dt.Rows.Add(row);
            }

            return dt;
        }

        public static DataTable ToDataTable(this IXLWorksheet worksheet)
        {
            var dt = new DataTable();
            dt.TableName = worksheet.Name;
            var firstRow = true;
            foreach (var row in worksheet.Rows())
            {
                //Use the first row to add columns to DataTable.
                if (firstRow)
                {
                    foreach (var cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }

                    firstRow = false;
                }
                else
                {
                    //Add rows to DataTable.
                    dt.Rows.Add();
                    var i = 0;

                    foreach (var cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber,
                                 row.LastCellUsed().Address.ColumnNumber))
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }

            return dt;
        }

        public static XLWorkbook ToExcel(this IEnumerable<DataTable> tables)
        {
            var workbook = new XLWorkbook();
            foreach (var dataTable in tables)
            {
                var worksheet = workbook.AddWorksheet(dataTable, dataTable.TableName);
                worksheet.Columns().AdjustToContents(15.0, 50.0);
            }

            return workbook;
        }
    }
}