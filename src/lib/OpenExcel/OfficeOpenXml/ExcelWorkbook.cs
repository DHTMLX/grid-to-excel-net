using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel.OfficeOpenXml
{
    public class ExcelWorkbook
    {
        public ExcelDocument Document { get; protected set; }

        public ExcelWorksheets Worksheets { get; protected set; }

        internal ExcelWorkbook(ExcelDocument parent)
        {
            this.Document = parent;
            this.Worksheets = new ExcelWorksheets(parent);
            

        }
    }
}
