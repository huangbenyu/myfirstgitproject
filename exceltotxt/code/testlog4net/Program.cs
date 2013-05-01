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

        private static readonly ILog logger =
       LogManager.GetLogger(typeof(Program));


        static void Main(string[] args)
        {


            BasicConfigurator.Configure();

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
