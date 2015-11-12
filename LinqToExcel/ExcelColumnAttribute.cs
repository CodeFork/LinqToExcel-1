using System;

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
