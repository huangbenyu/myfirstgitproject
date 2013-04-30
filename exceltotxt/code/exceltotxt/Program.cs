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

            if (test.LoadExcelFile())
            {
                int result = test.GetContent(sheetName);

                test.Close();

                utf8configdata = test.outstring;
                return result;

            }
            utf8configdata = "";
            return -1;
          

        }
        static public int SaveExcelData(string filename, string exceldata)
        {
            if (File.Exists(filename))
            {
                StreamReader sr = new StreamReader(filename);
                string olddata = sr.ReadToEnd();

                sr.Close();

                if (exceldata == olddata)
                {              
                    return 1;
                }
            }
          

            {
                StreamWriter sw;
                sw = new StreamWriter(filename, false);
                sw.Write(exceldata);
                sw.Close();
            }

            return 0;

        }
        
        static void Main(string[] args)
        {


            string filename = "./config/designdata.xml";

            XmlDocument srcFile = new XmlDocument();

            srcFile.Load(filename);

            ExcelConfigFile config = new ExcelConfigFile();
            config.Parse(srcFile);


            Console.WriteLine("start !!!");
            foreach (ExcelData exceldata in config.excelDatalist)
            {

                string outdata;

                string srcfilename = Directory.GetCurrentDirectory()+"/data/" + exceldata.GetFileName() + ".xlsx";
                if (File.Exists(srcfilename))
                {

                    GetExcelData(srcfilename, exceldata.GetSheetName(), out outdata);

                    string outfilename = Directory.GetCurrentDirectory() + "/data/temp/" + exceldata.GetFileName() + "." + exceldata.GetSheetName() + ".txt";

                    SaveExcelData(outfilename, outdata);
                }

            }

            //StreamWriter sw;
            //sw = new StreamWriter("./data/test.sheet1.txt" , false);
            //sw.WriteLine(outdata);
            //sw.Close();


            ////Excel.ReadExcelSheet("Book1.xlsx", "sheet1");

            //ExcelHelper test = new ExcelHelper("H:/workspace/myfirstgitproject/exceltotxt/Bin/data/test.xlsx");


            ////Array testlist = 
            // int result =test.GetContent("Sheet1");

            // test.Close();

            ////foreach (string  a  in testlist)
            ////{
            ////    Console.WriteLine(a);
            ////}

            Console.WriteLine("Copy file to Server and client directory");

            foreach (ExcelData exceldata in config.excelDatalist)
            {

     

                string srcfilename = Directory.GetCurrentDirectory() + "/data/temp/" + exceldata.GetFileName() + "." + exceldata.GetSheetName() + ".txt";

                if (File.Exists(srcfilename))
                {

                    string serverfile = Directory.GetCurrentDirectory() + config.ServerPath + exceldata.GetFileName() + "." + exceldata.GetSheetName() + ".txt";

                    File.Copy(srcfilename, serverfile,true);

                    string clientfile = Directory.GetCurrentDirectory() + config.ClientPath + exceldata.GetFileName() + "." + exceldata.GetSheetName() + ".txt";


                    File.Copy(srcfilename, clientfile,true);

             
                }

            }


            Console.WriteLine("Finish  !!!");

        }
    }
}
