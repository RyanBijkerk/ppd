using Parse_Performance_Data.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Parse_Performance_Data
{
    class FileItemHandler
    {
        public void Check(string file)
        {
            var error = new ErrorHandler();
            if (!File.Exists(file))
            {
                Console.WriteLine("{0}: Cannot find file: {1}", DateTime.Now, file);
                error.Exit(2);
            }
        }

        public List<Results> Parse(string file, List<string> metrics, char seperator)
        {
            var error = new ErrorHandler();
            
            // Opening csv file
            try
            {
                // Check if file can be read
                var tempFileReader = new StreamReader(file);
                tempFileReader.ReadLine();
                tempFileReader.Close();
                
            }
            catch (Exception)
            {
                Console.WriteLine("{0}: Cannot open file: {1}", DateTime.Now, file);
                error.Exit(4);
            }

            // Interval
            var offset = Offset(file, seperator);

            // File read
            var fileReader = new StreamReader(file);

            // Boolen to adjust the array
            var createArray = true;
            var totalResults = new string[0, 0];
            var objectTotalResults = new List<Results>();

            // Create results list
            var fileResults = new List<string>();

            // Adding all headers
            var fileHeader = fileReader.ReadLine().Split(seperator);
            var j = 0;

            foreach (var metric in metrics)
            {
                // Verify each header
                for (int i = 0; i <= fileHeader.Count() - 1; i++)
                {
                    // Check if header containts the metric
                    if (fileHeader[i].Contains(metric))
                    {
                        // Add header to set of results
                        fileResults.Add(metric);
                        var objectFileResult = new Results
                        {
                            File = file,
                            Header = metric
                        };
 
                        var dataResults = new List<string>();

                        TimeSpan timeStamp = new TimeSpan(0, 0, 0);

                        // Gatheing results of selected metric
                        while (!fileReader.EndOfStream)
                        {
                            var csvLine = fileReader.ReadLine();
                            var csvValues = csvLine.Split(seperator);

                            // Check for columns to prevent out of bounds exception
                            if (i <= csvValues.Count() - 1)
                            {
                                // Adding results to list
                                if (csvValues[i] != fileHeader[i])
                                {
                                    fileResults.Add(csvValues[i].Trim(new Char[] { '"' }));
                                    dataResults.Add(csvValues[i].Trim(new Char[] { '"' }));

                                    var value = 0.0;
                                    try
                                    {
                                        value = Convert.ToDouble(csvValues[i].Trim(new Char[] { '"' }));
                                    }
                                    catch (Exception)
                                    {
                                        //Console.WriteLine(e);
                                    }

                                    var interval = timeStamp;
                                    var dateTime = Convert.ToDateTime(csvValues[0].Trim(new Char[] { '"' }));

                                    timeStamp = timeStamp.Add(TimeSpan.FromSeconds(offset.Time));
                                }
                            }
                        }

                        // Create array if not exsits (know the size of multi array)
                        if (createArray == true)
                        {
                            totalResults = new string[metrics.Count(), fileResults.Count()];
                            createArray = false;
                        }

                        // set counter for mutli array
                        var k = 0;

                        // add result to object and list with all the results
                        objectFileResult.Data = dataResults;
                        objectTotalResults.Add(objectFileResult);

                        // Add data to total array
                        foreach (var item in fileResults.ToArray())
                        {
                            totalResults[j, k] = item;
                            k = k + 1;
                        }

                        // Next metric in the array
                        j = j + 1;

                        // Clear data and set reader position to first result
                        fileResults.Clear();
                        fileReader.BaseStream.Position = 1;

                        // Data found exit for loop
                        break;
                    }
                }
            }
            // Closing open csv file
            fileReader.Close();

            // Returing all results of the metric
            return objectTotalResults;
        }

        public Offset Offset(string file, char separator)
        {
            // Opening csv file
            var fileReader = new StreamReader(file);

            // Adding all headers
            var fileHeader = fileReader.ReadLine().Split(separator);

            // Define list and default time
            var timeList = new List<DateTime>();
            var timeDiffList = new List<double>();
            var time = new double();

            // Read all lines and add to timeList
            while (!fileReader.EndOfStream)
            {
                var csvLine = fileReader.ReadLine().Split(separator);
                var timeLine = csvLine[0].Substring(csvLine[0].LastIndexOf(' ') + 1).TrimEnd('\"', '\\');
                timeList.Add(DateTime.Parse(timeLine));
            }

            // Get rounded times
            for (int i = 0; i < timeList.Count; i++)
            {
                if (i + 1 <= timeList.Count - 1)
                {
                    var num = Math.Floor((timeList[i + 1] - timeList[i]).TotalSeconds);
                    var round = ((int)Math.Round(num / 10.0)) * 10;
                    timeDiffList.Add(round);
                }
            }

            var uniques = timeDiffList.Distinct().ToList();

            if (uniques.Count > 1)
            {
                var bestCount = 0;

                foreach (var unique in uniques)
                {

                    var amountDiff = timeDiffList.Count(x => x.Equals(unique));
                    if (amountDiff > bestCount)
                    {
                        bestCount = amountDiff;
                        time = unique;
                    }
                }


                //Console.WriteLine("Not good");
            }
            else
            {
                time = timeDiffList[0];
            }

            var retunObject = new Offset { Lines = timeList.Count, Time = time };

            // Closing open csv file
            fileReader.Close();

            // Returing all results of the metric
            return retunObject;
        }

        public double Time(string file)
        {
            // Opening csv file
            var fileReader = new StreamReader(file);

            // Adding all headers
            var fileHeader = fileReader.ReadLine().Split(',');
            var fileFirstLine = fileReader.ReadLine().Split(',');
            var fileSecondLine = fileReader.ReadLine().Split(',');

            // Getting only the time stamps
            var timeStamp1 = fileFirstLine[0];
            timeStamp1 = timeStamp1.Substring(timeStamp1.LastIndexOf(' ') + 1).TrimEnd('\"', '\\');

            var timeStamp2 = fileSecondLine[0];
            timeStamp2 = timeStamp2.Substring(timeStamp2.LastIndexOf(' ') + 1).TrimEnd('\"', '\\');

            // Parse for time
            var stamp1 = DateTime.Parse(timeStamp1);
            var stamp2 = DateTime.Parse(timeStamp2);

            // Rounding off number
            var timeDiff = Math.Floor((stamp2 - stamp1).TotalSeconds);

            // Closing open csv file
            fileReader.Close();

            // Returing all results of the metric
            return timeDiff;
        }
    }
}
