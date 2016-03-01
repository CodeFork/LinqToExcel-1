using System;

namespace ExcelToLinq
{
    /// <summary>
    /// Attribute to map a property to a column field in the excel file
    /// </summary>
    public class ExcelColumnAttribute : Attribute
    {
        public ExcelColumnAttribute()
        {

        }
        /// <summary>
        /// Attribute to map a property to a column field in the excel file
        /// </summary>
        /// <param name="Name">Name of the column in the excel file</param>
        public ExcelColumnAttribute(string Name)
        {
            this.Name = Name;
        }
        public string Name { get; set; }
    }
}
