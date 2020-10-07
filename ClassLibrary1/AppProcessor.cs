using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic
{
    public static class AppProcessor
    {
        public static void Run()
        {
            Application excelApp = new Application();

            Workbook workbook = excelApp.Workbooks.Open(@"C:\Users\pawel.iwanowski\Desktop\test.xlsx");
            Worksheet worksheet = workbook.Worksheets[1];

            worksheet.Cells[1, "A"].Value = "test";

            workbook.SaveAs(@"C:\Users\pawel.iwanowski\Desktop\test2.xlsx");
            workbook.Close();

        }
    }
}
