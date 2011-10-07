using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;


using DHTMLX.Export.Excel;

namespace Grid2Excel.Controllers
{
    [HandleError]
    public class GeneratorController : Controller
    {
        [HttpPost, ValidateInput(false)]
        public ActionResult Generate()
        {
            var generator = new ExcelWriter();
            var xml = this.Request.Form["grid_xml"];
            xml = this.Server.UrlDecode(xml);
            var stream = generator.Generate(xml);
            return File(stream.ToArray(), generator.ContentType, "grid.xlsx");          
        }

        
    }
}
