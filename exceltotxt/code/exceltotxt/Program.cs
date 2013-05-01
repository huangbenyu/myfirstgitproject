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


        static public int GetExcelData(string excelFileName, ExcelData exceldata)
        {
            ExcelHelper test = new ExcelHelper(excelFileName);

            if (test.LoadExcelFile())
            {
                foreach (Sheetdata sheetdata in exceldata.sheetDatalist)
                {
                    test.ClearData();

                    int result = test.GetContent(sheetdata.sheetName);

                    if (0 == result)
                    {
                        string outfilename = Directory.GetCurrentDirectory() + "/data/temp/" + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

                        SaveExcelData(outfilename, test.outstring);
                    }

                }

                test.Close();
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

            
                string srcfilename = Directory.GetCurrentDirectory()+"/data/" + exceldata.GetFileName() + ".xlsx";

                Console.WriteLine("Read Excel File : {0}", srcfilename);

                if (File.Exists(srcfilename))
                {

                    GetExcelData(srcfilename, exceldata);
                }

            }

            Console.WriteLine("Copy file to Server and client directory");

            foreach (ExcelData exceldata in config.excelDatalist)
            {
                foreach (Sheetdata sheetdata in exceldata.sheetDatalist)
                {
                    string srcfilename = Directory.GetCurrentDirectory() + "/data/temp/" + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

                    if (File.Exists(srcfilename))
                    {
                        string serverfile = Directory.GetCurrentDirectory() + config.ServerPath + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

                        File.Copy(srcfilename, serverfile, true);
                        if (sheetdata.client)
                        {
                            string clientfile = Directory.GetCurrentDirectory() + config.ClientPath + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

                            File.Copy(srcfilename, clientfile, true);
                        }
                    }
                }

            }


            Console.WriteLine("Finish  !!!");

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(true);

        }
    }
}
