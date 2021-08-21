using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel.Npoi.Extension;
using NPOI.HSSF.Extractor;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace Excel.Npoi
{
    public class ExcelHelper
    {
        /// <summary>
        /// excel 转换成 dataTable
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="firstRowIsTitle"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static DataTable ConvertToDataTable(string filePath, bool firstRowIsTitle = true, int sheetIndex = 0)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件不存在");
            }

            using var fs = File.OpenRead(filePath);
            return ConvertToDataTable(fs, firstRowIsTitle, sheetIndex);
        }

        /// <summary>
        /// 将excel转换为List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static List<T> ConvertToList<T>(string filePath, int sheetIndex = 0) where T : new()
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件不存在");
            }
            using var fs = File.OpenRead(filePath);
            return ConvertToList<T>(fs, sheetIndex);

        }
        /// <summary>
        /// excel 转换成 dataTable
        /// </summary>
        /// <param name="stream">excel流 </param>
        /// <param name="firstRowIsTitle"></param>
        /// <param name="sheetIndex">需要解析的sheet,默认不传是0</param>
        /// <returns></returns>
        public static DataTable ConvertToDataTable(Stream stream, bool firstRowIsTitle = true, int sheetIndex = 0)
        {
            var fileName = Path.GetExtension((stream as FileStream)?.Name);
            if (fileName is null)
                throw new ArgumentException("请检查传入文件流");
            stream.Position = 0;
            var fileNameExtension = fileName.ToLowerInvariant();
            if (fileNameExtension.Equals(".csv") || fileNameExtension.Equals(".tsv"))
            {
                return CommonHelper.CsvHandler(stream);
            }
            IWorkbook workbook = fileNameExtension switch
            {
                ".xls" => new HSSFWorkbook(stream),
                ".xlsx" => new XSSFWorkbook(stream),
                _ => throw new ArgumentException("不支持的格式")
            };
            return CommonHelper.CommonFormatHandler(workbook.GetSheetAt(sheetIndex), firstRowIsTitle);
        }


        /// <summary>
        /// 将excel转换为List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static List<T> ConvertToList<T>(Stream stream, int sheetIndex = 0) where T : new()
        {
            var fileName = Path.GetExtension((stream as FileStream)?.Name);
            if (fileName is null)
                throw new ArgumentException("请检查传入文件流");
            stream.Position = 0;
            var fileNameExtension = fileName.ToLowerInvariant();
            IWorkbook workbook = fileNameExtension switch
            {
                ".xls" => new HSSFWorkbook(stream),
                _ => new XSSFWorkbook(stream)
            };
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet is null || sheet.PhysicalNumberOfRows <= 0)
                return new List<T>();
            var entities = new List<T>(sheet.LastRowNum - 1);
            var propertyInfos = typeof(T).GetProperties();
            Dictionary<string, string> propertiesMatch = new();
            foreach (var property in propertyInfos)
            {
                var propertyAttribute = property.CustomAttributes
                    .Where(x => x.AttributeType.Name == "DescriptionAttribute").Select(y => new
                    {
                        property = property.Name,
                        correspondColumnName = y.ConstructorArguments.LastOrDefault().Value.ToString()
                    }).FirstOrDefault();
                if (propertyAttribute is null) continue;
                if (propertiesMatch.ContainsKey(propertyAttribute.property))
                {
                    propertiesMatch[propertyAttribute.property] = propertyAttribute.correspondColumnName;
                }
                else
                {
                    propertiesMatch.Add(propertyAttribute.correspondColumnName, propertyAttribute.property);
                }
            }
            IRow header = sheet.GetRow(sheet.FirstRowNum);
            List<Tuple<int, string>> list = new List<Tuple<int, string>>();
            var count = header.Cells.Max(x => x.ColumnIndex);
            for (int j = 0; j <= count; j++)
            {
                var val = header.Cells.Where(x => x.ColumnIndex == j)?.FirstOrDefault()?.StringCellValue;
                list.Add(Tuple.Create(j, val));
            }
            //数据  
            for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                T entity = new T();
                var type = typeof(T);
                for (int j = 0; j <= count; j++)
                {
                    //列名
                    var headCell = list[j].Item2;
                    if (headCell == null)
                    {
                        continue;
                    }
                    //值
                    var valueType = CommonHelper.GetValueType(sheet.GetRow(i).GetCell(j));
                    //空格赋值
                    if (valueType == null || valueType.ToString() == string.Empty) continue;
                    var judgeColumnInClass = CommonHelper.JudgeColumnInClass(propertiesMatch, propertyInfos, headCell);
                    if (!string.IsNullOrWhiteSpace(judgeColumnInClass))
                    {
                        type.GetProperty(judgeColumnInClass)?.SetValue(entity, valueType.ToString(), null);
                    }
                }
                entities.Add(entity);
            }
            return entities;
        }


        public static void ConvertListToExcelStream<T>(IEnumerable<T> entities, Stream stream)
        {
            var dataTable = CommonHelper.EntitiesToDataTable(entities);
            var fs = stream as FileStream;
            if (fs is null)
            {
                throw new ArgumentException("错误的文件流");
            }
            var extension = Path.GetExtension(fs.Name);
            IWorkbook workbook = extension.ToLowerInvariant() switch
            {
                ".xlsx" => new XSSFWorkbook(),
                ".xls" => new HSSFWorkbook(),
                _ => throw new ArgumentException("不支持的excel格式")
            };
            CommonHelper.DataTableToExcel(dataTable, stream, workbook);
        }
    }
}
