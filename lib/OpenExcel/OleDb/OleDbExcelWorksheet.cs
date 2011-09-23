using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace OpenExcel.OleDb
{
    public class OleDbExcelWorksheet
    {
        private WeakReference _cachedTable;
        private WeakReference _cachedTableIMEX;

        public string Name { get; protected set; }
        public OleDbExcelReader Reader { get; protected set; }
        public OleDbCells Cells { get; protected set; }

        internal OleDbExcelWorksheet(string name, OleDbExcelReader parent)
        {
            this.Name = name;
            this.Reader = parent;
            this.Cells = new OleDbCells(this);
        }

        internal DataTable GetCachedTable()
        {
            DataTable dt;
            if (_cachedTable == null || (dt = (DataTable)_cachedTable.Target) == null)
            {
                using (OleDbConnection conn = this.Reader.OpenConnection(false))
                {
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM ['" + this.Name.Replace("'", "''") + "$']", conn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    dt = new DataTable();
                    da.Fill(dt);
                    _cachedTable = new WeakReference(dt);
                }
            }
            return dt;
        }

        internal DataTable GetCachedTableIMEX()
        {
            DataTable dt;
            if (_cachedTableIMEX == null || (dt = (DataTable)_cachedTableIMEX.Target) == null)
            {
                using (OleDbConnection conn = this.Reader.OpenConnection(true))
                {
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM ['" + this.Name.Replace("'", "''") + "$']", conn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    dt = new DataTable();
                    da.Fill(dt);
                    _cachedTableIMEX = new WeakReference(dt);
                }
            }
            return dt;
        }
    }
}
