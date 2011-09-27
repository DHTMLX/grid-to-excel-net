using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenExcel.Common;

namespace OpenExcel.OleDb
{
    public class OleDbCells
    {
        public OleDbExcelWorksheet Worksheet { get; protected set; }

        public OleDbCells(OleDbExcelWorksheet wsheet)
        {
            this.Worksheet = wsheet;
        }

        public OleDbCell this[string address]
        {
            get
            {
                RowColumn rc = ExcelAddress.ToRowColumn(address);
                return this[rc.Row, rc.Column];
            }
        }

        public OleDbCell this[uint row, uint col]
        {
            get
            {
                return new OleDbCell(row, col, this.Worksheet);
            }
        }
    }
}
