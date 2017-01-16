using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace TrxConverter
{
    public static class Program
    {
        private static ILog Logger { get; } = LogManager.GetLogger(typeof(Program));

        public static void Main(string[] args)
        {
            TrxLogger.Setup();

            var options = new TrxCliParams();
            
            var bContinue = true;

            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {   
                
                var trxConversion = new TrxConversion(options);
                try
                {
                    bContinue = trxConversion.Convert();
                }
                catch (Exception e)
                {
                    Logger.Error("Conversion failed. ", e);
                    throw;
                }
                
            }
            else
            {
                // Display the default usage information
                Console.WriteLine(options.GetUsage());
            }

            if (bContinue)
            {
                PrintSuccess();
            }
        }

        private static void PrintSuccess()
        {
            Logger.Info("Successful conversion of TRX files.");
        }
    }
}
