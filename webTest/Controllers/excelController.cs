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


        public ActionResult create()
        {
            return View();
        }

        [HttpPost]
        public ActionResult create(string fileName)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                string path = Server.MapPath(fileName);
                DataTable data = createTmpData();

                cExcel.excel ex = new cExcel.excel();
                ViewBag.excelCreateFlag = ex.createExcel(path, data, "sheet_test");
            }
            
            return View();
        }

        private DataTable createTmpData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id", typeof(int));
            dt.Columns.Add("name", typeof(string));
            dt.Columns.Add("lastname", typeof(string));
            dt.Columns.Add("gender", typeof(string));
            dt.Columns.Add("date", typeof(DateTime));
            dt.Columns.Add("date2", typeof(DateTime));
            dt.Columns.Add("remark", typeof(string));

            DataRow dr;
            for (int i = 0; i < 5; i++)
            {
                dr = dt.NewRow();
                dr["id"] = (i + 1);
                dr["name"] = string.Format("test_{0}", i);
                dr["lastname"] = string.Format("lname_{0}", i);
                dr["date"] = DateTime.Now.AddDays(i);
                if (i == 2) dr["date2"] = DateTime.Now;
                dr["remark"] = Guid.NewGuid().ToString();
                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}
