using NPOI.HSSF.Extractor;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace Excel.Npoi.NetCore.Tests
{
    public class Student
    {
        public string Name { get; set; }
        [Description("Gender1")]
        public string Gender { get; set; }
    }
    public class UnitTest1
    {
        [Fact]
        public void ToDataTableTest()
        {
            using (var stream = new FileStream(@".\ExcelFiles\ListToExcelTest.xlsx", FileMode.Open))
            {
                var convertToDataTable = ExcelHelper.ConvertToDataTable(stream);
                //��dataTable
                for (var i = 0; i < convertToDataTable.Rows.Count; i++)
                {
                    var data = convertToDataTable.Rows[i];
                    if (data.ItemArray.Count() <= -1)
                    {
                        continue;
                    }
                    data.ItemArray[0]?.ToString();
                }
                Assert.NotNull(convertToDataTable);
            }
        }

        /// <summary>
        /// excel to list ����
        /// </summary>
        [Theory]
        [InlineData(@".\ExcelFiles\ListToExcelTest.xlsx")]
        public void ToListTest<T>(string path)
        {
            using (var fileStream = File.OpenRead(path))
            {
                var type = typeof(User);
                var convertToList = ExcelHelper.ConvertToList<User>(fileStream);
                Assert.True(convertToList.Count == 999);
            }
        }

        [Fact]

        public void GenerateExcelTest()
        {
            List<Student> students = new();
            Random rd = new Random();
            for (int i = 0; i < 1000; i++)
            {
                students.Add(new Student()
                {
                    Gender = rd.Next(2) == 0 ? "男" : "女",
                    Name = rd.Next(2) == 0 ? "123" : "456",
                });
            }

            using var fs = File.Create(@".\ExcelFiles\123123.xlsx");
            Excel.Npoi.ExcelHelper.ConvertListToExcelStream(students, fs);
        }

    }
}
