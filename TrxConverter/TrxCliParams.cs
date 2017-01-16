using System.Collections.Generic;
using System.IO;
using System.Text;
using CommandLine;
using CommandLine.Text;

namespace TrxConverter
{
    public class TrxCliParams
    {
        [Option('d', "directory", Required = true, HelpText = "Directory of input files to be processed.")]
        public string TrxFileDir { get; set; }


        [Option('e', "excel", Required = true, HelpText = "Output excel file of processed files.")]
        public string OutputExcelFile { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this);
        }
    }
}