using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;


using grid_excel_net;

namespace Grid2Excel.Controllers
{
    [HandleError]
    public class GeneratorController : Controller
    {
        
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
