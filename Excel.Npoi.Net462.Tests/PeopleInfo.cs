using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Npoi.Net462.Tests
{
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
}
