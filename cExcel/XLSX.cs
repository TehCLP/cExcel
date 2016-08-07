using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace cExcel
{
    internal class XLSX
    {
        public DataTable readToDT(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                DataTable dt = readAndCreateDataTable(fi, "string");
                return dt;
            }
            else return null;
        }
        public DataTable readToDT(Stream file)
        {
            DataTable dt = readAndCreateDataTable(file, "stream");
            return dt;
        }
        private DataTable readAndCreateDataTable(object file, string type)
        {
            try
            {
                DataTable dt = null;

                dynamic fi;
                if (type == "string")
                    fi = file as FileInfo;
                else
                    fi = file as Stream;

                using (ExcelPackage package = new ExcelPackage(fi))
                {
                    using (ExcelWorkbook workBook = package.Workbook)
                    {
                        if (workBook != null)
                        {
                            if (workBook.Worksheets.Count > 0)
                            {
                                using (ExcelWorksheet oSheet = workBook.Worksheets.First())
                                {
                                    int totalRows = oSheet.Dimension.End.Row;
                                    int totalCols = oSheet.Dimension.End.Column;

                                    dt = new DataTable(oSheet.Name);
                                    DataRow dr = null;

                                    for (int i = 1; i <= totalRows; i++)
                                    {
                                        if (i > 1) dr = dt.Rows.Add();
                                        for (int j = 1; j <= totalCols; j++)
                                        {
                                            if (i == 1)
                                                dt.Columns.Add(oSheet.Cells[i, j].Value.ToString());
                                            else
                                                dr[j - 1] = oSheet.Cells[i, j].Value.ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                return null;
            }
        }
    }
}
