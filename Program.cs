using System;
using System.Diagnostics;
using System.Reflection;
using Parse_Performance_Data.Code;

namespace Parse_Performance_Data
{
    class Program
    {

        static void Main(string[] args)
            {
            
            // Set version
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            
            Console.WriteLine("");
            Console.WriteLine("PPD - Parse Performance Data v{0}", version);
            Console.WriteLine("Copyright (C) 2010-2016 Ryan Bijkerk");
            Console.WriteLine("Logit Blog - www.logitblog.com");
            Console.WriteLine("");

            var seperator = ',';

            //ExcelCheck.Version();

            // Check the arguments
            var argumentHandler = new ArgumentHandler();
            var arguments = argumentHandler.Parse(args);
            
            // Collecting the metrics
            var settingsMetric = new MetricsHandler();
            var metrics = settingsMetric.Load();

            // File parsing
            var file = new FileItemHandler();
            var report = new ReportHandler();

            // Start for the timer
            var startTime = DateTime.Now;

            foreach (var fileItem in arguments.fileLocation)
            {
                // File Check
                file.Check(fileItem);

                Console.WriteLine("{0}: Working on file: {1}", DateTime.Now, fileItem);

                // File Parse for the data
                var fileResults = file.Parse(fileItem, metrics, seperator);

                // File time interval data
                var fileOffset = file.Offset(fileItem, seperator);

                // Create Excel file
                var reportFile = report.Create(fileItem);

                Console.WriteLine("{0}: Creating Excel report on location: {1}", DateTime.Now, reportFile);

                // Add data to Excel file
                report.AddData(reportFile, fileResults, fileOffset.Time, fileOffset.Lines);
                // Add charts to Excel file
                report.AddCharts(reportFile);
            }

            // Done and reporting total time
            var stopTime = DateTime.Now;
            Console.WriteLine("{0}: Done in {1} sec! ", DateTime.Now, Math.Round(stopTime.Subtract(startTime).TotalSeconds));
        }
    }
}
