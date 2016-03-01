using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ExcelToLinq;

namespace MVC.Test.Excel2Linq.Example.Models
{
    public class Person
    {
        /// <summary>
        /// Example of a property that does not exist in the excel file
        /// </summary>
        [ExcelNotMap]
        public int Id { get; set; }
        /// <summary>
        /// Map the property 'Name' with the Excel Column 'Name'
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Map the property 'LastName' with the Excel Column 'Last Name'
        /// </summary>
        public string LastName { get; set; }
        /// <summary>
        /// Map the property 'PhoneNumber' with the Excel Column 'Phone'
        /// </summary>
        [ExcelColumn("Phone")]
        public string PhoneNumber { get; set; }
    }
}