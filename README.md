## Excel操作类库

### Nuget 搜索 `Excel.Npoi` 并选择安装.

引入命名空间 `using Excel.Npoi`

#### 方法1：ExcelHelper.ConvertToDataTable

将Excel 转成DataTable,支持xls,xlsx ,csv,tsv

参数信息：`Stream stream(文件流), bool firstRowIsTitle = true(第一行是否标题), int sheetIndex (xls或xlsx格式第几个sheet)= 0`

测试代码:Excel.Npoi.Net462.Tests\ExcelTest.cs\ExcelToDataTableTest

```c#
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
            }
			也可以不用传stream 直接 var convertToDataTable = ExcelHelper.ConvertToDataTable(filePath, false);
```

#### 方法2:ExcelHelper.ConvertToList<T>

将Excel转换成List<T> ,**如需Excel列表题与实体对应** 需在property上加上 **特性** **Description**    支持Excel格式 xls xlsx

参数信息: `Stream stream(文件流), int sheetIndex() = 0`

测试代码:Excel.Npoi.Net462.Tests\ExcelTest.cs\ExcelToListTest

```c#

			//path:excel文件路径
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
            public class PeopleInfo
            {
                [Description("First Name")]
                public string FirstName { get; set; }
                [Description("Last Name")]
                public string LastName { get; set; }
                public string Gender { get; set; }
                public string Country { get; set; }
                public string Age { get; set; }
                public string Date { get; set; }
                public string Id { get; set; }
                [Description("LikeNess")]
                public string Like { get; set; }
            }
也可以不用传stream 直接 var convertToList = ExcelHelper.ConvertToDataTable(filePath, false);
```

#### 方法3  ConvertListToExcelStream 

根据List<Entity>生成excel，支持xlsx,xls

参数信息: `IEnumerable<T> entities(实体集合), Stream stream(文件流)`

测试代码:Excel.Npoi.Net462.Tests\ExcelTest.cs\ListToExcelTest

```c#
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
```

