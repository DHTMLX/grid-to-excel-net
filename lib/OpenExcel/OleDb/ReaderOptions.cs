using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel.OleDb
{
    public class ReaderOptions
    {
        public IMEXOptions IMEX { get; set; }

        public ReaderOptions()
        {
            this.IMEX = IMEXOptions.No;
        }
    }
}
