using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using log4net;
using log4net.Config;

namespace testlog4net
{
    class Program
    {

        public static readonly ILog logger =
       LogManager.GetLogger("server");


        static void Main(string[] args)
        {

			XmlConfigurator.Configure(new System.IO.FileInfo("log4net.xml"));

           // BasicConfigurator.Configure();

            logger.Debug("Here is a debug log.");
            logger.Info("... and an Info log.");
            logger.Warn("... and a warning.");
            logger.Error("... and an error.");
            logger.Fatal("... and a fatal error.");



            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(true);
        }
    }
}
