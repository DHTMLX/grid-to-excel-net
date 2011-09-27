using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel.OfficeOpenXml.Style
{
    public interface IStylable
    {
        ExcelStyle Style { get; set; }
    }
}
