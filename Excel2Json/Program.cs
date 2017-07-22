using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

namespace Excel2Json
{
    class Program
    {
        const int rowCount = 5;
        const int colCount = 5;

        static void Main(string[] args)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\mobinology\Desktop\tester.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Create a list to hold the content to be read
            List<string> lines = new List<string>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                //Create a string to hold the content of this row
                string line = "";

                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        line += xlRange.Cells[i, j].Value2.ToString();
                    }
                }

                lines.Add(line);
            }

            //write lines to a txt file
            System.IO.File.WriteAllLines(@"C:\Users\mobinology\Desktop\tester.txt", lines);

            //Console.Read();
        }
    }
}

