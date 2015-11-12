using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLinq
{
   public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute()
        {

        }
        public ExcelColumnAttribute(string Name)
        {
            this.Name = Name;
        }
        public string Name { get; set; }
    }
}
