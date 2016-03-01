using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel;

namespace ExcelToLinq
{
    /// <summary>
    /// Main Class that converts a Excel file to IEnumerable of the specified entity class
    /// </summary>
    public class ExcelToEntity : IDisposable
    {
        /// <summary>
        /// The excel table has row header with the information
        /// </summary>
        public bool HasHeaders { get; set; }
        /// <summary>
        /// Excel Data Reader Interface to convert excel to data table
        /// </summary>
        IExcelDataReader edr;
        /// <summary>
        /// Manual contructor by File stream and excel version type
        /// </summary>
        /// <param name="s">Can be a memory stream or file stream or byte array</param>
        /// <param name="et">Excel type of the file given (binary xls or open format xlsx)</param>
        public ExcelToEntity(Stream s, ExcelType et)
        {
            HasHeaders = true;
            switch (et)
            {
                case ExcelType.xlsx:
                    edr = ExcelReaderFactory.CreateOpenXmlReader(s);
                    break;
                case ExcelType.xls:
                    edr = ExcelReaderFactory.CreateBinaryReader(s);
                    break;
                default:
                    break;
            }
            edr.IsFirstRowAsColumnNames = false;
        }
        /// <summary>
        /// Automatic constructor based on html form post multipart input type file
        /// </summary>
        /// <param name="File">posted file from asp html request or MVC post action</param>
        public ExcelToEntity(System.Web.HttpPostedFileBase File)
        {
            if (File == null) throw new ArgumentNullException("HttpPostedFileBase File is null");

            HasHeaders = true;
            if (File.FileName.Trim().ToLower().EndsWith(".xlsx"))
                edr = ExcelReaderFactory.CreateOpenXmlReader(File.InputStream);
            else
                edr = ExcelReaderFactory.CreateBinaryReader(File.InputStream);

            edr.IsFirstRowAsColumnNames = false;
        }
        /// <summary>
        /// Read the data table obtain from the converted excel file and transform to the entity class
        /// </summary>
        /// <typeparam name="T">Entity Class to transform to</typeparam>
        /// <param name="StartRow">Start row from the data table</param>
        /// <param name="StartColumn">Start column from the data table</param>
        /// <param name="SheetName">sheet name where to read the data</param>
        /// <returns></returns>
        public IEnumerable<T> Read<T>(int StartRow = 1, int StartColumn = 1, string SheetName = null) where T : new()
        {
            var dt = getTable(StartRow, StartColumn, SheetName);
            if (!HasHeaders) StartRow--;
            for (int i = StartRow; i < dt.Rows.Count; i++)
            {
                var r = dt.Rows[i];
                if (!(r.ItemArray.All(x => string.IsNullOrWhiteSpace(x.ToString()))))
                    yield return GetEntity<T>(r);
            }
        }
        /// <summary>
        /// internal method to transform the excel file to a data table
        /// </summary>
        /// <param name="startRow">Start row from the data table</param>
        /// <param name="startCol">Start column from the data table</param>
        /// <param name="sheetName">sheet name where to read the data</param>
        /// <returns></returns>
        private System.Data.DataTable getTable(int startRow, int startCol, string sheetName)
        {
            var ds = edr.AsDataSet();
            System.Data.DataTable dt;
            if (!string.IsNullOrEmpty(sheetName))
            {
                dt = ds.Tables[sheetName];
            }
            else
            {
                dt = ds.Tables[0];
            }
            if (HasHeaders)
            {
                var drHeader = dt.Rows[startRow - 1];
                for (int i = startCol - 1; i < dt.Columns.Count; i++)
                {
                    var name = drHeader[i];
                    if (name != null && name is string)
                    {
                        var cname = new String(((string)name).Where(c => char.IsLetterOrDigit(c)).ToArray());
                        if (char.IsDigit(cname.FirstOrDefault())) cname = "_" + cname;
                        if (!string.IsNullOrWhiteSpace(cname))
                        {
                            try
                            {
                                dt.Columns[i].ColumnName = cname;
                            }
                            catch (System.Data.DuplicateNameException)
                            {
                                dt.Columns[i].ColumnName = cname + "_" + i.ToString();
                            }
                        }
                    };
                }
            }
            return dt;
        }
        /// <summary>
        /// internal method to convert a data row to a data entity
        /// </summary>
        /// <typeparam name="t">Entity Class</typeparam>
        /// <param name="row">Data row to grab the data from</param>
        /// <returns></returns>
        private t GetEntity<t>(System.Data.DataRow row) where t : new()
        {
            var entity = new t();
            var properties = typeof(t).GetProperties();

            foreach (var property in properties.Where(x => x.GetCustomAttributes(typeof(ExcelNotMapAttribute), true).SingleOrDefault() == null))
            {
                //Get the description attribute
                var CustomAttribute = (ExcelColumnAttribute)property.GetCustomAttributes(typeof(ExcelColumnAttribute), true).SingleOrDefault();
                string propName = CustomAttribute == null ? property.Name : CustomAttribute.Name;

                object v = null;
                try { v = row[propName]; }
                catch (Exception) { continue; }

                try
                {
                    property.SetValue(entity, v.GetType() == property.PropertyType ? v :
                        Convert.ChangeType(row[propName], property.PropertyType));
                }
                catch (Exception)
                {
                    throw new InvalidCastException("Unable to cast value '" + v.ToString() +
                        "' to type '" + property.PropertyType.Name + "' for property: " + property.Name +
                        " row: " + row.ItemArray.Aggregate((a, b) => a.ToString() + "," + b.ToString()));
                }
            }

            return entity;
        }
        /// <summary>
        /// implement the garbage colector and close the unmanaged excel to data table reader
        /// </summary>
        public void Dispose()
        {
            if (!edr.IsClosed)
            {
                edr.Close();
            }
        }
    }
}
