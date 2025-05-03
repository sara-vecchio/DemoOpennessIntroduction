using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DemoOpennessIntroduction
{
    public static class ReadExcel
    {
       // public int nRow;
       // public int sInput;
        
        public static void ReadCell(Excel.Range xlr, out string sInput, int nColumn, int nRow)
        {
            try 
            {
                sInput = xlr.Cells[nRow, nColumn].Value2.ToString(); 
            }
            catch 
            {
                sInput = null; 
            }
        }

        public static int ReadRange(Excel.Range xlr)
        {
            int nTotRow = 1;
            try 
            {
                nTotRow = xlr.Rows.Count;
            }
            catch 
            {
               
            }
            return nTotRow;
        }


    }
}
