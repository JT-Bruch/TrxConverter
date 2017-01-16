using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using log4net;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TrxConverter
{
    public class TrxConversion
    {
        private static ILog Logger { get; } = LogManager.GetLogger(typeof(TrxConversion));
        private TrxCliParams CliOptions { get; set; }

        public TrxConversion(TrxCliParams options)
        {
            CliOptions = options;
        }

        public bool Convert()
        {
            List<TrxFile> trxFiles;
            IEnumerable<TrxInfo> trxInfos = null;


            var bContinue = ConvertCliFiles(GetFilesFromDirectory(CliOptions.TrxFileDir), out trxFiles);

            if (bContinue && trxFiles != null)
            {
                bContinue = ConvertTrxFilesToDatalist(trxFiles, out trxInfos);
            }

            if (bContinue && trxInfos != null)
            {
                bContinue = ConvertDatalistToExcel(trxInfos, CliOptions.OutputExcelFile);
            }

            return bContinue;
        }

        private IEnumerable<string> GetFilesFromDirectory(string cliOptionsTrxFileDir)
        {
            IEnumerable<string> fileList = new List<string>();
            if (Directory.Exists(cliOptionsTrxFileDir))
            {
                fileList = Directory.EnumerateFiles(cliOptionsTrxFileDir, "*.trx");
            }
            return fileList;
        }


        private static bool ConvertCliFiles(IEnumerable<string> cliInputs, out List<TrxFile> trxFiles)
        {
            trxFiles = new List<TrxFile>();
            try
            {
                foreach (var filePath in cliInputs)
                {
                    if (File.Exists(filePath))
                    {
                        trxFiles.Add(new TrxFile(filePath));
                    }
                    else
                    {
                        Logger.Debug($"File {filePath} does not exist");
                    }

                }
            }
            catch (Exception e)
            {
                Logger.Error(e);
                Console.WriteLine(e);
                return false;
            }
           
            return true;
        }

        private static bool ConvertTrxFilesToDatalist(IEnumerable<TrxFile> trxFiles, out IEnumerable<TrxInfo> trxInfos )
        {
            trxInfos = new List<TrxInfo>();
            try
            {
                foreach (var trx in trxFiles)
                {
                    var trxFile = XDocument.Load(trx.GetFilePath());
                    if (trxFile.Root == null) continue;

                    var df = trxFile.Root.Name.Namespace;
                    trxInfos = trxInfos.Concat((from ut in trxFile.Descendants(df + "UnitTestResult")
                                select new TrxInfo()
                                {
                                    ExecutionId = (string)ut.Attribute("executionId"),
                                    TestId = (string)ut.Attribute("testId"),
                                    TestName = (string)ut.Attribute("testName"),
                                    Outcome = (string)ut.Attribute("outcome"),
                                    Duration = (string)ut.Attribute("duration"),
                                    StdOut =  GetChildElementValue(ut, "StdOut"),
                                    ErrorMessage = GetChildElementValue(ut, "Message"),
                                    StackTrace = GetChildElementValue(ut, "StackTrace"),
                                }));

                }
            }
            catch (Exception e)
            {
                Logger.Error(e);
                Console.WriteLine(e);
                return false;
            }
            
            return true;
        }

        private static string GetChildElementValue(XElement unitTest, string valToGet)
        {
            var retVal = "";
            if (unitTest.HasElements)
            {
                retVal = unitTest.Descendants().SingleOrDefault(x => x.Name.LocalName == valToGet)?.Value;
            }
            return retVal;
        }

        private bool ConvertDatalistToExcel(IEnumerable<TrxInfo> trxInfos, string cliOptionsOutputExcelFile)
        {
            
            try
            {
                if (CheckForOutputFile(cliOptionsOutputExcelFile))
                {
                    using (var excelDoc = new ExcelPackage(new FileInfo(cliOptionsOutputExcelFile)))
                    {
                        //Here setting some document properties
                        excelDoc.Workbook.Properties.Author = "Crownpeak Technologies";
                        excelDoc.Workbook.Properties.Title = "Unit Test Data Mining";

                        //Create a sheet
                        excelDoc.Workbook.Worksheets.Add("Test Sheet");
                        ExcelWorksheet ws = excelDoc.Workbook.Worksheets[1];
                        ws.Name = "Test Sheet"; //Setting Sheet's name
                        ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                        ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

                        //Merging cells and create a center heading for out table
                        ws.Cells[1, 1].Value = "Unit Test Data";
                        ws.Cells[1, 1, 1, typeof(TrxInfo).GetProperties().Length].Merge = true;
                        ws.Cells[1, 1, 1, typeof(TrxInfo).GetProperties().Length].Style.Font.Bold = true;
                        ws.Cells[1, 1, 1, typeof(TrxInfo).GetProperties().Length].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        var colIndex = 1;
                        const int rowIndex = 2;

                        foreach (var propInfo in typeof(TrxInfo).GetProperties()) //Creating Headings
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            //Setting Value in cell
                            cell.Value = propInfo.Name;
                            colIndex++;
                        }


                        var dataRange = ws.Cells[3, 1];
                        dataRange.LoadFromCollection(trxInfos);
                        excelDoc.Save();

                    }
                }
                
            }
            catch (Exception e)
            {
                Logger.Error(e);
                Console.WriteLine(e);
                return false;
            }
            return true;
        }

        private bool CheckForOutputFile(string cliOptionsOutputExcelFile)
        {
            try
            {
                if (File.Exists(cliOptionsOutputExcelFile))
                {
                    Logger.Error("File already exists in this location!");
                    Logger.Error("Deleting File.");
                    File.Delete(cliOptionsOutputExcelFile);
                }
            }
            catch (Exception e)
            {
                Logger.Error("Failed to delete file.", e);
                Console.WriteLine(e);
                return false;
            }
            return true;
        }

    }
}