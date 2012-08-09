using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Reflection;

namespace exceltotxt
{
    class Excel
    {

        public static void ReadExcelSheet(string fileName, string sheetName)
        {
            if (!File.Exists(fileName))
            {
                Console.WriteLine("fileName %s not exist", fileName);
            }

            try
            {



                Microsoft.Office.Interop.Excel.ApplicationClass excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();

                Microsoft.Office.Interop.Excel.Workbooks wbs = excelApp.Workbooks;


                object objOpt = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(fileName, objOpt, false, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);

                //Microsoft.Office.Interop.Excel.Workbook wb = wbs.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                int count = wb.Worksheets.Count;

                string[] names = new string[count];

                for (int i = 1; i <= count; i++)
                {

                }

            }
            catch 
            {
                throw new Exception("未安装Excel或者未安装Excel对.net的编程支持");
            }

        }
    }
}
