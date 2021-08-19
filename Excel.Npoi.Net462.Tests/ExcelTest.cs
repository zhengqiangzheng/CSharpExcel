using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel.Npoi.Extension;
using Xunit;

namespace Excel.Npoi.Net462.Tests
{
    public class ExcelTest
    {
        /// <summary>
        /// 读csv tsv测试
        /// </summary>
        /// <param name="path"></param>
        [Theory]
        [InlineData(@".\ExcelFiles\csvtest3.csv")]
        [InlineData(@".\ExcelFiles\tsvtest1.tsv")]
        [InlineData(@".\ExcelFiles\csvtest1.csv")]
        [InlineData(@".\ExcelFiles\csvtest2.csv")]
        public void ReadCsvTest(string path)
        {
            using (var fs = File.OpenRead(path))
            {
                var dt = ExcelHelper.ConvertToDataTable(fs);
                //读dataTable
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    var data = dt.Rows[i];
                    if (data.ItemArray.Count() <= -1)
                    {
                        continue;
                    }
                    var 第i行第1列 = data.ItemArray[0]?.ToString();
                    var 第i行第2列 = data.ItemArray[1]?.ToString();
                    var 第i行第3列 = data.ItemArray[2]?.ToString();
                }
                Assert.NotNull(dt);
            }
        }
        /// <summary>
        /// 读xls xlsx测试
        /// </summary>
        [Theory]
        [InlineData(@".\ExcelFiles\Empty.xlsx")]
        [InlineData(@".\ExcelFiles\file_example_XLSX_5000.xlsx")]
        [InlineData(@".\ExcelFiles\file_example_XLS_5000.xls")]
        [InlineData(@".\ExcelFiles\file_example_XLS_1000.xls")]
        [InlineData(@".\ExcelFiles\file_example_XLSX_1000.xlsx")]
        public void ExcelToDataTableTest(string path)
        {
            using (var fs = File.OpenRead(path))
            {
                //第二个参数不传默认第一行是标题不读 传false读第一行
                var convertToDataTable = ExcelHelper.ConvertToDataTable(fs, false);
                //读dataTable
                for (var i = 0; i < convertToDataTable.Rows.Count; i++)
                {
                    var data = convertToDataTable.Rows[i];
                    if (data.ItemArray.Count() <= -1)
                    {
                        continue;
                    }
                    var 第i行第1列 = data.ItemArray[0]?.ToString();
                    var 第i行第2列 = data.ItemArray[1]?.ToString();
                    var 第i行第3列 = data.ItemArray[2]?.ToString();
                    var 第i行第4列 = data.ItemArray[3]?.ToString();
                    var 第i行第5列 = data.ItemArray[4]?.ToString();
                }
                Assert.NotNull(convertToDataTable);
            }
        }

        /// <summary>
        /// 读Excel生成List<T>测试
        /// </summary>
        /// <param name="path"></param>
        [Theory]
        // [InlineData(@".\ExcelFiles\test1.xlsx")]
        [InlineData(@".\ExcelFiles\file_example_XLSX_5000.xlsx")]
        [InlineData(@".\ExcelFiles\file_example_XLS_5000.xls")]
        [InlineData(@".\ExcelFiles\file_example_XLS_1000.xls")]
        [InlineData(@".\ExcelFiles\file_example_XLSX_1000.xlsx")]
        public void ExcelToListTest(string path)
        {
            using (var fs = File.OpenRead(path))
            {
                try
                {
                    var convertToList = ExcelHelper.ConvertToList<PeopleInfo>(fs);

                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }
        }


        /// <summary>
        /// 根据List生成Excel
        /// </summary>
        [Fact]
        public void ListToExcelTest()
        {
            var userEntities = new List<UserEntity>() { };
            var rd = new Random();
            for (int i = 1; i < 1000; i++)
            {

                rd.Next(2);
                userEntities.Add(new UserEntity()
                {
                    UserName = "asdfdfasd",
                    UserAddress = "lalsdflafadasdf",
                    UserId = i.ToString(),
                    UserPhone = "131231234",
                    UserPwd = rd.Next(2) == 0 ? null : "123444123541234"
                });
            }
            using (var fileStream = new FileStream(@".\ListToExcelTest.xlsx", FileMode.Create))
            {
                ExcelHelper.ConvertListToExcelStream(userEntities, fileStream);
            }
        }
    }
}
