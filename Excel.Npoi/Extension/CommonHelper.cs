using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel.Npoi.Extension
{
    public class CommonHelper
    {
        public static void DataTableToExcel(DataTable dataTable, Stream stream, IWorkbook workbook, ExcelFormat excelFormat = ExcelFormat.Xlsx, string sheetName = "sheet1")
        {
            ISheet excelSheet = workbook.CreateSheet(sheetName);
            List<string> columns = new List<string>();
            IRow row = excelSheet.CreateRow(0);
            int columnIndex = 0;

            foreach (DataColumn column in dataTable.Columns)
            {
                columns.Add(column.ColumnName);
                row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                columnIndex++;
            }

            int rowIndex = 1;
            foreach (DataRow dsrow in dataTable.Rows)
            {
                row = excelSheet.CreateRow(rowIndex);
                int cellIndex = 0;
                foreach (String col in columns)
                {
                    var o = dsrow[col];
                    var success = int.TryParse(o.ToString(), out var number);
                    if (success)
                    {
                        row.CreateCell(cellIndex).SetCellValue(number);
                    }
                    else
                    {
                        row.CreateCell(cellIndex).SetCellValue(o.ToString());

                    }
                    cellIndex++;
                }
                rowIndex++;
            }
            workbook.Write(stream);
        }

        /// <summary>
        /// List转换为DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="items"></param>
        /// <returns></returns>
        public static DataTable EntitiesToDataTable<T>(IEnumerable<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            Dictionary<string, string> propertiesMatch = new();

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var prop in props)
            {
                var propertyAttribute = prop.CustomAttributes
                    .Where(x => x.AttributeType.Name == "DescriptionAttribute").Select(y => new
                    {
                        property = prop.Name,
                        correspondColumnName = y.ConstructorArguments.LastOrDefault().Value.ToString()
                    }).FirstOrDefault();
                if (propertyAttribute != null)
                {
                    propertiesMatch.Add(propertyAttribute.property, propertyAttribute.correspondColumnName);
                }
            }
            foreach (PropertyInfo prop in props)
            {
                dataTable.Columns.Add(propertiesMatch.ContainsKey(prop.Name) ? propertiesMatch[prop.Name] : prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
        internal static DataTable CsvHandler(Stream stream)
        {
            var sr = new StreamReader(stream);
            DataTable dtTable = new DataTable();
            string[] headers = sr.ReadLine()?.Split(new char[] { ',', '\t' });
            foreach (string hd in headers)
            {
                dtTable.Columns.Add(hd);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine() ?? string.Empty, "(\t|,)(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)").Where(x => x != "," && x != "\t").ToArray();
                DataRow dr = dtTable.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                dtTable.Rows.Add(dr);
            }
            return dtTable;
        }
        internal static string JudgeColumnInClass(Dictionary<string, string> propertiesMatch, PropertyInfo[] propertyInfos,
            string headerCell)
        {
            var orDefault = propertiesMatch.Keys.FirstOrDefault(x => x.Equals(headerCell, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(orDefault))
            {
                return propertiesMatch[orDefault];
            }
            var firstOrDefault = propertyInfos.FirstOrDefault(x => x.Name.Equals(headerCell, StringComparison.OrdinalIgnoreCase));
            return firstOrDefault != null ? firstOrDefault.Name : "";
        }

        internal static DataTable CommonFormatHandler(ISheet sheet, bool firstRowIsTitle)
        {
            DataTable dtTable = new DataTable();
            //表头  
            IRow header = sheet.GetRow(sheet.FirstRowNum);
            if (header is null)
                throw new ArgumentException("Excel内容为空");
            List<int> columns = new List<int>();
            for (int i = 0; i < header.LastCellNum; i++)
            {
                object obj = GetValueType(header.GetCell(i));
                if (obj == null || obj.ToString() == string.Empty)
                {
                    dtTable.Columns.Add(new DataColumn("Columns" + i.ToString()));
                }
                else
                    dtTable.Columns.Add(new DataColumn(obj.ToString()));

                columns.Add(i);
            }
            //数据
            var startRowIndex = sheet.FirstRowNum + 1;
            if (!firstRowIsTitle)
            {
                startRowIndex = sheet.FirstRowNum;
            }
            for (int i = startRowIndex; i <= sheet.LastRowNum; i++)
            {
                DataRow dr = dtTable.NewRow();
                bool hasValue = false;
                foreach (int j in columns)
                {
                    dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                    if (dr[j] != null && dr[j].ToString() != string.Empty)
                    {
                        hasValue = true;
                    }
                }
                if (hasValue)
                {
                    dtTable.Rows.Add(dr);
                }
            }
            return dtTable;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        internal static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
