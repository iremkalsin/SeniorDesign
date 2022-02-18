using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorDesignImplementation
{
    public class ExcelDataWriter
    {
        IWorkbook workbook = new XSSFWorkbook();
        public void WriteToSheet(string sheetName, double[,] array)
        {
            ISheet sheet = workbook.CreateSheet(sheetName);
            IRow row0 = sheet.CreateRow(0);
            ICell cell0 = row0.CreateCell(0);
            cell0.SetCellValue(sheetName);
            for (int x = 1; x < array.GetLength(1); x++)
            {
                ICell cell = row0.CreateCell(x);
                cell.SetCellValue(x);
            }
            for (int i = 1; i <= array.GetLength(0); i++)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell = sheet.GetRow(i).CreateCell(0);
                cell.SetCellValue(i);
                for (int j = 1; j < array.GetLength(1); j++)
                {
                    ICell cell2 = row.CreateCell(j);
                    cell2.SetCellValue(array[i-1, j]);
                }
            }
        }

        public void Save()
        {
            using (var fs = File.Create($"result3.xlsx"))
            {
                workbook.Write(fs);
            }
        }
    }
}
