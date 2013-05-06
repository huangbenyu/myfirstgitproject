using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;

using System.IO;

using log4net;
using log4net.Config;


namespace Exceltotxt
{
    class Program
    {
		public static readonly ILog Logger = LogManager.GetLogger("server");

        static public int SaveExcelData(string filename, string exceldata)
        {
			Logger.DebugFormat("SaveExcelData :{0}", filename);
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
			DBExcelHelper test = new DBExcelHelper(excelFileName);

            if (test.LoadExcelFile())
            {
				Logger.InfoFormat("Read  Excel File Name :{0}" ,excelFileName );
				foreach (Sheetdata sheetdata in exceldata.sheetDatalist)
                {
                    test.ClearData();

                    int result = test.GetContent(sheetdata.sheetName);

					Logger.InfoFormat("Get  Sheet Content :{0}", sheetdata.sheetName);

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
			
			XmlConfigurator.Configure(new System.IO.FileInfo("log4net.xml"));



			System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string filename = "./config/designdata.xml";

            XmlDocument srcFile = new XmlDocument();

            srcFile.Load(filename);

            ExcelConfigFile config = new ExcelConfigFile();
            config.Parse(srcFile);


            Console.WriteLine("start !!!");

    
            foreach (ExcelData exceldata in config.excelDatalist)
            {

            
                string srcfilename = Directory.GetCurrentDirectory()+"/data/" + exceldata.GetFileName() + ".xlsx";

                Logger.InfoFormat("Load Excel File : {0}", srcfilename);

                if (File.Exists(srcfilename))
                {

                    GetExcelData(srcfilename, exceldata);
                }

            }

			Logger.Info("Copy file to Server and client directory");

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


			Logger.Info("Finish  !!!");

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(true);

        }
    }
}
