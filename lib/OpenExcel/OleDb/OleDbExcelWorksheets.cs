using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace OpenExcel.OleDb
{
    public class OleDbExcelWorksheets : IEnumerable<OleDbExcelWorksheet>
    {
        public OleDbExcelReader Reader { get; protected set; }

        internal OleDbExcelWorksheets(OleDbExcelReader parent)
        {
            this.Reader = parent;
        }

        public OleDbExcelWorksheet this[string name]
        {
            get
            {
                return new OleDbExcelWorksheet(name, this.Reader);
            }
        }

        #region IEnumerable<OleDbWorksheet> Members

        public IEnumerator<OleDbExcelWorksheet> GetEnumerator()
        {
            foreach (var i in EnumerateWorksheets())
                yield return i;
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            foreach (var i in EnumerateWorksheets())
                yield return i;
        }

        private IEnumerable<OleDbExcelWorksheet> EnumerateWorksheets()
        {
            using (var conn = this.Reader.OpenConnection(false))
            {
                DataTable tblSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                var worksheetNames = from r in tblSchema.Rows.Cast<DataRow>()
                                     let tableName = (string)r["TABLE_NAME"]
                                     where tableName.EndsWith("$") ||
                                           (tableName.StartsWith("'") && tableName.EndsWith("$'"))
                                     select tableName;
                foreach (string name in worksheetNames)
                {
                    string tableName = name;
                    // Remove quotes and "$"
                    if (name.StartsWith("'") && name.EndsWith("$'"))
                        tableName = name.Substring(1, name.Length - 3).Replace("''","'");
                    else
                        tableName = name.Substring(0, name.Length - 1);
                    yield return new OleDbExcelWorksheet(tableName, this.Reader);
                }
            }
        }

        #endregion
    }
}
