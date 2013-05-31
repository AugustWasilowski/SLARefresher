using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace SLARefresher
{
    class Program
    {
        private static int TotalFileCount;
        private static int CompletedFileCount;
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static StringBuilder trace = new StringBuilder();
        private static StringBuilder errors = new StringBuilder();
        private static bool HasErrors;

        static void Main(string[] args)
        {
            HasErrors = false;
            TotalFileCount = 0;
            CompletedFileCount = 0;
            DateTime Start = DateTime.Now;
            logger.Trace("Refreshing Excel files in " + ConfigurationManager.AppSettings["RootDirectoryToSearch"]);
            trace.Append("Refreshing Excel files in " + ConfigurationManager.AppSettings["RootDirectoryToSearch"]);
            trace.AppendLine();
            Console.WriteLine("Press Ctrl-C to abort.");
            
            RefreshFiles(ConfigurationManager.AppSettings["RootDirectoryToSearch"]);
            
            logger.Info("SLA Refresh has completed. " 
                + CompletedFileCount.ToString() 
                + " of " + TotalFileCount + " files located in " 
                + ConfigurationManager.AppSettings["RootDirectoryToSearch"] 
                + " were refreshed in: " + (DateTime.Now - Start).ToString()
                + "\r\n"
                + "\r\n"
                + "See trace below: "
                + "\r\n"
                + "\r\n"
                + trace.ToString()                
                );

            if (HasErrors)
            {
                logger.Error(errors);
            }

            if (Boolean.Parse(ConfigurationManager.AppSettings["WaitForUserInteractionBeforeClosing"]))
            {
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey(true);    
            }            
        }
        /// <summary>
        /// Traverses a Directory and all sub directories, finding files to open and refresh.
        /// </summary>
        /// <param name="pathToTraverse">Root Directory that should contain the Excel files</param>
        private static void RefreshFiles(string pathToTraverse)
        {
            try
            {
                foreach (var file in new DirectoryInfo(pathToTraverse).GetFiles("*.xlsx", SearchOption.AllDirectories))
                {
                    TotalFileCount++;
                    string filePath = file.Directory + "\\" + file;
                    // Process File
                    if (file.ToString().Substring(0, 1) != "~") // Don't process files that are Temp Files opened already by Excel
                    {
                        DateTime StartFile = DateTime.Now;
                        logger.Trace("Refreshing File: " + filePath);
                        trace.Append("Refreshing File: " + filePath);
                        trace.AppendLine();
                        if (Boolean.Parse(ConfigurationManager.AppSettings["RefreshFiles"]))
                        {

                            RefreshExcelWorkbook(filePath);    
                        }

                        logger.Trace("Completed in: " + (DateTime.Now - StartFile).ToString());
                        trace.Append("Completed in: " + (DateTime.Now - StartFile).ToString());
                        trace.AppendLine();
                        Console.ForegroundColor = ConsoleColor.White;
                        logger.Trace("--------------------------------------------------------------");
                        trace.Append("--------------------------------------------------------------");
                        trace.AppendLine();
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }
                }
            }
            catch (Exception ex)
            {                
                Console.ForegroundColor = ConsoleColor.Red;
                HasErrors = true;
                errors.Append(ex.Message);
                errors.AppendLine();
                trace.Append(ex.Message);
                trace.AppendLine();
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }

        /// <summary>  
        /// Refreshes an excel workbooks data then saves the refreshed data to a new file  
        /// </summary>  
        /// <param name="workbookPath">  
        /// <returns>Path of file with refreshed data or null on failure</returns>  
        private static bool RefreshExcelWorkbook(string workbookPath)
        {
            try
            {
                Application excelApp = new Application();

                excelApp.DisplayAlerts = false;

                // if Visible is set to false we get an exception for some reason  
                excelApp.Visible = true;

                Workbook excelWorkbook = excelApp.Workbooks.Open(
                     workbookPath,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value,
                     System.Reflection.Missing.Value);

                // pivot tables are normally refreshed in a background thread  
                // setting this to false means that we can save synchronously  
                foreach (PivotCache cache in excelWorkbook.PivotCaches())
                {
                    cache.BackgroundQuery = false;
                }

                excelWorkbook.RefreshAll();
                excelWorkbook.Save();
                
                if (Boolean.Parse(ConfigurationManager.AppSettings["CopyToAlternateLocation"]))
                {
                    try
                    {
                        var refreshedWorkBookPath = workbookPath.Replace(ConfigurationManager.AppSettings["RootDirectoryToSearch"], ConfigurationManager.AppSettings["AlternateLocation"]);
                        excelWorkbook.SaveCopyAs(refreshedWorkBookPath);
                        logger.Trace("Copy saved as: " + refreshedWorkBookPath);
                        trace.Append("Copy saved as: " + refreshedWorkBookPath);
                        trace.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        HasErrors = true;
                        errors.Append("Error saving copy: " + ex.Message);
                        errors.AppendLine();
                        trace.Append("Error saving copy: " + ex.Message);
                        trace.AppendLine();
                        Console.ForegroundColor = ConsoleColor.Gray;
                    }                    
                }

                // ensure all excel objects are closed and references are released  
                // if this is not done an instance of excel will stay in the running processes  
                excelWorkbook.Close(false, workbookPath, null);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                excelWorkbook = null;

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                CompletedFileCount++;

                Console.ForegroundColor = ConsoleColor.Green;
                logger.Trace("Sucessfully refreshed and saved: " + workbookPath);
                trace.Append("Sucessfully refreshed and saved: " + workbookPath);
                trace.AppendLine();
                Console.ForegroundColor = ConsoleColor.Gray;

                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                HasErrors = true;
                errors.Append("Error refreshing file: " + ex.Message);
                errors.AppendLine();
                trace.Append("Error refreshing file: " + ex.Message);
                trace.AppendLine();
                Console.ForegroundColor = ConsoleColor.Gray;

                return false;
            }
        }  
    }
}
