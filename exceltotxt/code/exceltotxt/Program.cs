using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace exceltotxt
{
    class Program
    {
        static void Main(string[] args)
        {

            //Excel.ReadExcelSheet("Book1.xlsx", "sheet1");

            ExcelHelper test = new ExcelHelper("H:/workspace/myfirstgitproject/exceltotxt/Bin/Book1.xlsx");


            //Array testlist = 
                test.GetContent("Sheet1");

            //foreach (string  a  in testlist)
            //{
            //    Console.WriteLine(a);
            //}

        }
    }
}
