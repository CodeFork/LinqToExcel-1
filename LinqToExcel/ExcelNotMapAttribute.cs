using System;

namespace ExcelToLinq
{
    /// <summary>
    /// Attribute to skip the mapping of a property in the selected entity class
    /// </summary>
    public class ExcelNotMapAttribute : Attribute
    {
        public ExcelNotMapAttribute()
        {

        }
    }
}
