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
        private int _startRow = 1;
        private int _startCol = 1;

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

        public bool createExcel(FileInfo fi, DataTable dt, string sheetName)
        {
            bool b = false;

            try
            {
                // Gen File Excel
                using (var xlsx = new ExcelPackage(fi))
                {
                    ExcelWorksheet sheet = xlsx.Workbook.Worksheets.Add(sheetName);
                    genTheadExcel(ref sheet, dt);

                    int row = _startRow + 1; // start rows at lineIndex 2
                    int totalCol = dt.Columns.Count;

                    foreach (DataRow dr in dt.Rows)
                    {
                        for (int col = 0; col < totalCol; col++)
                        {
                            sheet.Cells[row, (col + _startCol)].Value = dr[col];
                        }
                        row++;
                    }

                    xlsx.Save();
                    b = true;
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                b = false;
            }

            return b;
        }
        private void genTheadExcel(ref ExcelWorksheet sheet, DataTable dt)
        {
            int x = _startRow; // row
            int y = _startCol; // col
            foreach (DataColumn col in dt.Columns)
            {
                sheet.Cells[x, y].Value = col.ColumnName;
                y++;
            }

            // # set style
            using (var range = sheet.Cells[x, x, x, y])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                //range.Style.Font.Color.SetColor(Color.White);
                range.Style.ShrinkToFit = false;
            }
        }
    }
}
