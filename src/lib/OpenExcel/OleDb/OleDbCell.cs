using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using OpenExcel.Common;

namespace OpenExcel.OleDb
{
    public class OleDbCell
    {
        public OleDbExcelWorksheet Worksheet { get; protected set; }
        public uint Row { get; protected set; }
        public uint Column { get; protected set; }

        public string Address
        {
            get
            {
                return RowColumn.ToAddress(this.Row, this.Column);
            }
        }

        public OleDbCell(uint row, uint col, OleDbExcelWorksheet wsheet)
        {
            this.Row = row;
            this.Column = col;
            this.Worksheet = wsheet;
        }

        public object Value
        {
            get
            {
                Func<DataTable>[] sources = null;
                switch (this.Worksheet.Reader.Options.IMEX)
                {
                    case IMEXOptions.Yes:
                        sources = new Func<DataTable>[] {
                                        new Func<DataTable>(this.Worksheet.GetCachedTableIMEX)
                                  };
                        break;
                    case IMEXOptions.No:
                        sources = new Func<DataTable>[] {
                                        new Func<DataTable>(this.Worksheet.GetCachedTable)
                                  };
                        break;
                    case IMEXOptions.IMEXFirst:
                        sources = new Func<DataTable>[] {
                                        new Func<DataTable>(this.Worksheet.GetCachedTableIMEX),
                                        new Func<DataTable>(this.Worksheet.GetCachedTable),
                                  };
                        break;
                    case IMEXOptions.NoIMEXFirst:
                        sources = new Func<DataTable>[] {
                                        new Func<DataTable>(this.Worksheet.GetCachedTable),
                                        new Func<DataTable>(this.Worksheet.GetCachedTableIMEX),
                                  };
                        break;
                }

                foreach (Func<DataTable> src in sources)
                {
                    DataTable dt = src();
                    if (this.Row > dt.Rows.Count)
                        return null;
                    if (this.Column > dt.Columns.Count)
                        return null;

                    object val = dt.Rows[(int)this.Row - 1][(int)this.Column - 1];
                    if (val != DBNull.Value)
                        return val;
                }
                return DBNull.Value;
            }
        }
    }
}
