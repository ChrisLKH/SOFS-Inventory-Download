using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Linq;

namespace SOFSInventoryDownloader
{
    public static class ExcelHelper
    {
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pck = new ExcelPackage();

            using var stream = File.OpenRead(path);
            pck.Load(stream);

            var ws = pck.Workbook.Worksheets.First();
            var tbl = new DataTable();

            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");

            int startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                    row[cell.Start.Column - 1] = cell.Text;
                tbl.Rows.Add(row);
            }

            return tbl;
        }
    }
}
