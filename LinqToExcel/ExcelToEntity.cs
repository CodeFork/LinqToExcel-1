using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;

namespace ExcelToLinq
{
    public class ExcelToEntity : IDisposable
    {
        public bool HasHeaders { get; set; }

        IExcelDataReader edr;
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
                        if (!string.IsNullOrWhiteSpace(cname)) dt.Columns[i].ColumnName = cname;
                    };
                }
            }
            return dt;
        }

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

        public void Dispose()
        {
            if (!edr.IsClosed)
            {
                edr.Close();
            }
        }
    }
}
