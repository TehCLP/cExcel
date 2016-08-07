using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.IO;

namespace cExcel
{
    internal class XLS
    {
        public DataTable readToDT(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Exists)
            {
                DataTable dt = readAndCreateDataTable(path, "string");
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
                dynamic fi;
                if (type == "string")
                    fi = file as string;
                else
                    fi = file as Stream;

                Workbook book = Workbook.Load(fi);
                if (book.Worksheets.Count > 0)
                {
                    Worksheet sheet = book.Worksheets[0];

                    DataTable dt = new DataTable(sheet.Name);
                    DataRow dr = null;

                    for (int i = sheet.Cells.FirstRowIndex; i <= sheet.Cells.LastRowIndex; i++)
                    {
                        Row row = sheet.Cells.GetRow(i);
                        if (i > sheet.Cells.FirstRowIndex) dr = dt.Rows.Add();

                        for (int j = row.FirstColIndex; j <= row.LastColIndex; j++)
                        {
                            Cell cell = row.GetCell(j);
                            if (i == sheet.Cells.FirstRowIndex)
                                dt.Columns.Add(cell.StringValue);
                            else
                            {
                                if (cell.Format.FormatType == CellFormatType.Date || cell.Format.FormatType == CellFormatType.DateTime)
                                    dr[j - row.FirstColIndex] = cell.DateTimeValue;
                                else
                                    dr[j - row.FirstColIndex] = cell.StringValue;
                            }
                        }
                    }

                    return dt;
                }
                else return null;
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                return null;
            }
        }

        internal void create()
        {
            //create new xls file
            string file = "C:\\newdoc.xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("First Sheet");
            worksheet.Cells[0, 1] = new Cell((short)1);
            worksheet.Cells[2, 0] = new Cell(9999999);
            worksheet.Cells[3, 3] = new Cell((decimal)3.45);
            worksheet.Cells[2, 2] = new Cell("Text string");
            worksheet.Cells[2, 4] = new Cell("Second string");
            worksheet.Cells[4, 0] = new Cell(32764.5, "#,##0.00");
            worksheet.Cells[5, 1] = new Cell(DateTime.Now, @"YYYY\-MM\-DD");
            worksheet.Cells.ColumnWidth[0, 1] = 3000;
            workbook.Worksheets.Add(worksheet);
            workbook.Save(file);

        }
    }
}
