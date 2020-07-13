using System;
using System.Collections.Generic;
using System.IO;

namespace Parse_Performance_Data
{
    class MetricsHandler
    {
        public List<string> Load()
        {
            // Error variable created
            var error = new ErrorHandler();

            var metrics = new List<string>();
            var location = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var metricsFile = location + "\\Metrics.txt";

            if (!File.Exists(metricsFile))
            {
                Console.WriteLine("{0}: Cannot find: {1}", DateTime.Now, metricsFile);
                Console.WriteLine("{0}: Creating Metrics.txt file", DateTime.Now);

                // Creating file
                //File.Create(metricsFile);
                
                var exampleMetrics = "# VMware ESXtop example for CPU: (_Total)\\% Util Time";
                File.WriteAllText(metricsFile, exampleMetrics);
                
                // Return Exit code 2
                error.Exit(2);
            }
            else
            {
                // Try to read file
                try
                {
                    var metricReader = new StreamReader(metricsFile);
                    string metricLine;

                    while ((metricLine = metricReader.ReadLine()) != null)
                    {
                        // Adding metrics to the file
                        if (!metricLine.StartsWith("#"))
                        {
                            if (metricLine != "")
                            {
                                metrics.Add(metricLine);
                            }
                        }
                    } 
                }
                catch (Exception)
                {
                    // Catch cannot read the file because it is locked
                    Console.WriteLine("{0}: Cannot read file: {1}", DateTime.Now, metricsFile);
                    error.Exit(4);
                }
            }

            // Check if the list contain the metrics
            if (metrics.Count == 0)
            {
                Console.WriteLine("{0}: Metrics file does not contain any metrics", DateTime.Now);
                error.Exit(13);
            }
            
            return metrics;
        }
    }
}
