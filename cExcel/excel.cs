using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace cExcel
{
    public class excel
    {
        public DataTable readToDataTable(string path)
        {
            string extension = Path.GetExtension(path);
            DataTable dt = read(path, extension, "string");
            return dt;
        }
        public DataTable readToDataTable(Stream file, string fileName)
        {
            DataTable dt = read(file, fileName, "stream");
            return dt;
        }
        private DataTable read(object file, string extension, string type)
        {
            string ext = extension.ToLower();
            if (ext.EndsWith(".xls") || ext.EndsWith(".xlsx"))
            {
                try
                {
                    dynamic excel;
                    if (ext.EndsWith(".xls"))
                        excel = new XLS();
                    else
                        excel = new XLSX();

                    dynamic fi;
                    if (type == "string")
                        fi = file as string;
                    else
                        fi = file as Stream;

                    DataTable dt = new DataTable();
                    dt = excel.readToDT(fi);

                    return dt;
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
    }
}
