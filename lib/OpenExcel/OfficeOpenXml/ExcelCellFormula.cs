using OpenExcel.Common;
using OpenExcel.OfficeOpenXml.Internal;

namespace OpenExcel.OfficeOpenXml
{
    public class ExcelCellFormula
    {
        private uint _row;
        private uint _col;
        private ExcelWorksheet _wsheet;

        internal ExcelCellFormula(uint row, uint col, ExcelWorksheet w)
        {
            _row = row;
            _col = col;
            _wsheet = w;
        }

        public string Text
        {
            get
            {
                CellProxy c = _wsheet.GetCell(_row, _col);
                if (c != null)
                    if (c.Formula != null)
                        return c.Formula.Text;
                return null;
            }
            set
            {
                CellProxy c = _wsheet.EnsureCell(_row, _col);
                if (c.Formula == null)
                {
                    c.CreateFormula();
                    c.Formula.Text = value;
                }
            }
        }

        public void CopyTo(uint targetRow, uint targetCol)
        {
            string newFormulaText = ExcelFormula.Translate(this.Text, (int)targetRow - (int)_row, (int)targetCol - (int)_col);
            _wsheet.Cells[targetRow, targetCol].Formula.Text = newFormulaText;
        }

        public void CopyTo(string address)
        {
            RowColumn rc = ExcelAddress.ToRowColumn(address);
            CopyTo(rc.Row, rc.Column);
        }

        public void Remove()
        {
            CellProxy c = _wsheet.GetCell(_row, _col);
            if (c != null)
                c.RemoveFormula();
        }
    }
}
