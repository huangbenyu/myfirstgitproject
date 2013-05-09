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


        static public int GetExcelData(string excelFileName, ExcelData exceldata ,string excelpath)
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
                        string outfilename = Directory.GetCurrentDirectory() + excelpath + "temp/" + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

                        SaveExcelData(outfilename, test.outstring);
                    }

                }

                test.Close();
            }
            return 0;
        }

        static void Main(string[] args)
        {

			XmlConfigurator.Configure(new System.IO.FileInfo("./config/log4net.xml"));



			System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            string filename = "./config/designdata.xml";

            XmlDocument srcFile = new XmlDocument();

            srcFile.Load(filename);

            ExcelConfigFile config = new ExcelConfigFile();
            config.Parse(srcFile);



			string tempath = Directory.GetCurrentDirectory() + config.Excelpath + "temp/";
			if (!System.IO.Directory.Exists(tempath))
				//执行以下这条语句,就可以创建该文件夹了
				System.IO.Directory.CreateDirectory(tempath);

            Console.WriteLine("start !!!");

    
            foreach (ExcelData exceldata in config.excelDatalist)
            {


				string srcfilename = Directory.GetCurrentDirectory() + config.Excelpath + exceldata.GetFileName() + ".xlsx";

               // Logger.InfoFormat("Load Excel File : {0}", srcfilename);

                if (File.Exists(srcfilename))
                {
					GetExcelData(srcfilename, exceldata, config.Excelpath);
                }

            }

			Logger.Info("Copy file to Server and client directory");
			//检测目录是否存在
			
			string serverpaths = Directory.GetCurrentDirectory() + config.ServerPath;
			if (!System.IO.Directory.Exists(serverpaths))
				//执行以下这条语句,就可以创建该文件夹了
				System.IO.Directory.CreateDirectory(serverpaths);


			string clientpath = Directory.GetCurrentDirectory() + config.ClientPath;
			if (!System.IO.Directory.Exists(clientpath))
				//执行以下这条语句,就可以创建该文件夹了
				System.IO.Directory.CreateDirectory(clientpath);


            foreach (ExcelData exceldata in config.excelDatalist)
            {
                foreach (Sheetdata sheetdata in exceldata.sheetDatalist)
                {
					string srcfilename = Directory.GetCurrentDirectory() + config.Excelpath + "temp/" + exceldata.GetFileName() + "." + sheetdata.sheetName + ".txt";

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
