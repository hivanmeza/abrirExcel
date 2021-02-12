using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace tareaxls.clases
{
    class clsEstructura
    {
        public string nombre { get; set; }
        public string direccion { get; set; }

        public List<clsEstructura> cargaDatosXLS()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int rCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\casti\Desktop\tarea1 - copia.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            List<clsEstructura> todos = new List<clsEstructura>();
            clsEstructura individual = new clsEstructura();

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {

                individual.nombre = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                individual.direccion = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;

                todos.Add(individual);
                individual = new clsEstructura();
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            return todos;

        }



    }



}

