using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace PDFReader
{
    class ExcelReading
    {
        
        public static void Main(String [] args)
        {
            string path1 = "E:\\INTADM21DF006.xlsx";
            var excelFIle = new Application();
            Workbook wb = excelFIle.Workbooks.Open(path1);
            int sheetNo = 1;
            Worksheet ws = wb.Worksheets[sheetNo];
            Console.ReadLine();
        }
    }
}
