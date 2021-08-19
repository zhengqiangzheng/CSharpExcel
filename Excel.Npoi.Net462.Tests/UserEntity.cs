using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Npoi.Net462.Tests
{
    public class UserEntity
    {
        [Description("Newuserid")]
        public string UserId { get; set; }
        public string UserName { get; set; }
        public string UserPwd { get; set; }
        public string UserAddress { get; set; }
        public string UserPhone { get; set; }
        public string UserPhone2 { get; set; }
        public People People { get; set; }

    }

    public class People
    {
        public int Id { get; set; }
    }
}
