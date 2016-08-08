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
            if (TempData["excelReadModel"] != null)
            {
                ViewBag.excelReadModel = TempData["excelReadModel"];
            }

            return View();
        }

        [HttpPost]
        public ActionResult upload()
        {
            DataTable dt = new DataTable();

            string _path = Request["path"];
            string _fileName = Request["fileName"];
            
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

            TempData["excelReadModel"] = dt;

            if (!string.IsNullOrEmpty(_fileName))
            {
                return Json(true);
            }
            else
            {
                return RedirectToAction("read");
            }
        }

    }
}
