using System.Xml;

namespace DHTMLX.Export.Excel
{
    public class ExcelColumn
    {
        private string colName;
        private string type;
        private string align;
        private int colspan;
        private int rowspan;
        private int width = 0;
        private int height = 1;
        private bool is_footer = false;

        public void Parse(XmlElement parent)
        {
            is_footer = parent.ParentNode.ParentNode.Name.Equals("foot");
          
            if (parent.HasChildNodes)
                colName = parent.FirstChild.Value;
            else
                colName = "";

          
            if (parent.HasAttribute("width"))
            {
                width = int.Parse(parent.Attributes["width"].Value);
            }

            type = parent.GetAttribute("type");
            align = parent.GetAttribute("align");
        
            if (parent.HasAttribute("colspan"))
            {
                colspan = int.Parse(parent.Attributes["colspan"].Value);
            }
           
            if (parent.HasAttribute("rowspan"))
            {
                rowspan = int.Parse(parent.Attributes["rowspan"].Value);
            }
        }

        public int GetWidth()
        {
            return width;
        }

        public bool IsFooter()
        {
            return is_footer;
        }

        public void SetWidth(int width)
        {
            this.width = width;
        }

        public int GetColspan()
        {
            return colspan;
        }

        public int GetRowspan()
        {
            return rowspan;
        }

        public int GetHeight()
        {
            return height;
        }

        public void SetHeight(int height)
        {
            this.height = height;
        }

        public string GetName()
        {
            return colName;
        }

        public string getAlign()
        {
            return align;
        }

        public string getType()
        {
            return type;
        }
    }
}
