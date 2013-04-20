using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using System.Xml;

using System.IO;
namespace exceltotxt
{
    class Program
    {

        static  public int GetExcelData(string excelFileName, string sheetName, out string utf8configdata)
        {
             ExcelHelper test = new ExcelHelper(excelFileName);


            int result = test.GetContent(sheetName);

            test.Close();


            utf8configdata = test.outstring;

            return result;

        }
        
        static void Main(string[] args)
        {


            //string filename = "./config/designdata.xml";

            //XmlDocument srcFile = new XmlDocument();

            //srcFile.Load(filename);

            //ExcelConfigFile config = new ExcelConfigFile();
            //config.Parse(srcFile);


            string outdata;

            GetExcelData("H:/workspace/myfirstgitproject/exceltotxt/Bin/data/test.xlsx", "Sheet1",out outdata);


            StreamWriter sw;
            sw = new StreamWriter("./data/test.sheet1.txt" , false);
            sw.WriteLine(outdata);
            sw.Close();


            ////Excel.ReadExcelSheet("Book1.xlsx", "sheet1");

            //ExcelHelper test = new ExcelHelper("H:/workspace/myfirstgitproject/exceltotxt/Bin/data/test.xlsx");


            ////Array testlist = 
            // int result =test.GetContent("Sheet1");

            // test.Close();

            ////foreach (string  a  in testlist)
            ////{
            ////    Console.WriteLine(a);
            ////}

            Console.WriteLine("Finish  Content");

        }
    }
}
