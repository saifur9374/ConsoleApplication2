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
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open("c:\\Users\sarahma\Documents\Disney\Disney%20Destinations\HoneyComb\WHOLESALE%20-%20WDW%20-%20RAW%20DATA\WHOLESALE%20-%20WDW%20-%20RAW%20DATA/FY17\Wholesale%20Data_ADM_RN.xlsx"
");
            xlApp.Visible = true;
            foreach (Excel.Worksheet sht in xlWorkBook.Worksheets)
            {
                sht.Select();
                xlWorkBook.SaveAs(string.Format("{0}{1}.csv", "c:\\adfgetstarted\test", sht.Name), Excel.XlFileFormat.xlCSV, Excel.XlSaveAsAccessMode.xlNoChange);

            }
            xlWorkBook.Close(false);
        }
    }
}