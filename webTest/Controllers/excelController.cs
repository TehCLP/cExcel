using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using System.Data;

namespace webTest.Controllers
{
    public class excelController : Controller
    {
        public ActionResult read()
        {
            return View();
        }

        [HttpPost]
        public ActionResult read(DataTable model)
        {
            DataTable dt = new DataTable();

            string _path = Request["path"];

            if (!string.IsNullOrEmpty(_path))
            {
                _path = Server.MapPath(_path);

                cExcel.excel ex = new cExcel.excel();
                dt = ex.readToDataTable(_path);
            }
            else if (Request.Files.Count > 0)
            {
                var file = Request.Files[0];

                if (file != null && file.ContentLength > 0)
                {
                    var fileName = Path.GetFileName(file.FileName);
                    //var path = Path.Combine(Server.MapPath("~/files/"), fileName);
                    //file.SaveAs(path);

                    cExcel.excel ex = new cExcel.excel();
                    dt = ex.readToDataTable(file.InputStream, fileName);
                }
            }

            return View(dt);
        }

    }
}
