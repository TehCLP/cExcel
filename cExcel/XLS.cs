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
        private int _startRow = 1;
        private int _startCol = 1;

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

        public bool createExcel(FileInfo fi, DataTable dt, string sheetName)
        {
            bool b = false;

            try
            {
                //create(fi.FullName);

                Workbook xls = new Workbook();
                Worksheet sheet = new Worksheet(sheetName);
                genTheadExcel(ref sheet, dt);

                int row = _startRow + 1; // start rows at lineIndex 2
                int totalCol = dt.Columns.Count;

                foreach (DataRow dr in dt.Rows)
                {
                    for (int col = 0; col < totalCol; col++)
                    {
                        var type = dr[col].GetType();
                        if (dr[col] == System.DBNull.Value || dr[col] == null)
                        {
                            sheet.Cells[row, (col + _startCol)] = new Cell(dr[col].ToString());
                        }
                        else if (type == typeof(DateTime))
                        {
                            sheet.Cells[row, (col + _startCol)] = new Cell(dr[col], CellFormat.Date);
                        }
                        else
                        {
                            sheet.Cells[row, (col + _startCol)] = new Cell(dr[col]);
                        }
                    }
                    row++;
                }

                xls.Worksheets.Add(sheet);
                xls.Save(fi.FullName);

                b = true;
                
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                b = false;
            }

            return b;
        }
        private void genTheadExcel(ref Worksheet sheet, DataTable dt)
        {
            int x = _startRow; // row
            int y = _startCol; // col
            foreach (DataColumn col in dt.Columns)
            {
                sheet.Cells[x, y] = new Cell(col.ColumnName);
                y++;
            }
        }

        private void create(string path)
        {
            //create new xls file
            string _file = path; //"C:\\newdoc.xls";
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
            workbook.Save(_file);

        }
    }
}
