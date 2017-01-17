using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoAleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open("c:\\adfgetstarted\\testinput.xlsx");
            xlApp.Visible = true;
            foreach (Excel.Worksheet sht in xlWorkBook.Worksheets)
            {
                sht.Select();
                xlWorkBook.SaveAs(string.Format("{0}{1}.csv", "c:\\adfgetstarted\\testoutput", sht.Name), Excel.XlFileFormat.xlCSV, Excel.XlSaveAsAccessMode.xlNoChange);

            }
            xlWorkBook.Close(false);
        }
    }
}
