using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Parse_Performance_Data
{
    class ReportHandler
    {
        public void Version()
        {
            var error = new ErrorHandler();
            Type officeType14 = Type.GetTypeFromProgID("Excel.Application.14");
            Type officeType15 = Type.GetTypeFromProgID("Excel.Application.15");
            Type officeType16 = Type.GetTypeFromProgID("Excel.Application.16");
            if (officeType14 == null && officeType15 == null && officeType16 == null)
            {
                error.Exit(2013);
            }
        }

        public string Create(string file)
        {
            var error = new ErrorHandler();

            var cleanFile = Path.GetFileNameWithoutExtension(file);
            var location = Path.GetDirectoryName(file);
            var reportFile = location + "\\" + cleanFile + ".xlsx";

            if (!File.Exists(reportFile))
            {
                var excel = new Excel.Application { DisplayAlerts = false };
                // Creatring file
                excel.Workbooks.Add();

                // Creating sheet for Charts
                var workSheet = (Excel.Worksheet)excel.ActiveSheet;
                workSheet.Name = "Charts";

                // Delete default sheets in Office 2010
                for (int i = 1; i < excel.Sheets.Count +1; i++)
                {
                    if (excel.Sheets[i].Name != "Charts")
                    {
                        excel.Sheets[i].Delete();
                    }
                }

                // Save excel sheet
                workSheet.SaveAs(reportFile);

                // Close sheet
                excel.Quit();
            }
            else
            {
                Console.WriteLine("{0}: File already exsits: {1}", DateTime.Now, reportFile);
                error.Exit(183);
            }

            return reportFile;
        }

        public void AddData(string reportFile, string[,] results, double time)
        {
            var excel = new Excel.Application {DisplayAlerts = false};

            // Open report file
            excel.Workbooks.Open(reportFile);

            // Open the report Excel sheet
            Excel._Worksheet workSheet = (Excel.Worksheet)excel.ActiveSheet;

            // BLOCK EDIT
            var colomn = 0;

            var sheetNames = new List<string>();

            var arrayLength = results.GetLength(0) - 1;
            for (int i = 0; i <= arrayLength; i++)
            {
                var arraySecondLength = results.GetLength(1) - 1;
                for (int j = 0; j <= arraySecondLength; j++)
                {
                    if (j == 0 && results[i, j] != null)
                    {

                        // Set the sheetname based on metric from last /
                        var sheetName = results[i, j];
                        if (sheetName.Contains('\\'))
                        {
                            if (sheetName.Contains("(vmhba"))
                            {
                                sheetName = sheetName.Substring(sheetName.LastIndexOf('('));
                            }
                            else if (sheetName.Contains("vmnic"))
                            {
                                sheetName = sheetName.Substring(sheetName.LastIndexOf(":") + 1);
                                sheetName = "(" + sheetName;

                            }
                            else
                            {
                                sheetName = sheetName.Substring(sheetName.LastIndexOf('\\') + 1);
                            }

                        }
                        else if (sheetName.Contains(':'))
                        {
                            sheetName = sheetName.Substring(sheetName.LastIndexOf(':') + 1);
                        }

                        char[] charsToTrim = { '\\', '/', '?', '*', '[', ']', ':' };

                        // Replace each carh for a <space>
                        foreach (var trimChar in charsToTrim)
                        {
                            sheetName = sheetName.Replace(trimChar, ' ');
                        }

                        var maxLength = 30;
                        if (sheetName.Length > maxLength)
                        {
                            //var maxBeginPoint = sheetName.Length - maxLength;
                            sheetName = sheetName.Substring(0, maxLength);
                        }

                        if (!sheetNames.Exists(s => s.Equals(sheetName)))
                        {
                            sheetNames.Add(sheetName);
                        }
                        else
                        {
                            var sheetIndex = sheetNames.FindIndex(s => s.Equals(sheetName));
                            sheetName = sheetIndex.ToString() + sheetName;
                            sheetNames.Add(sheetName);
                        }
                        
                        // Add new sheet
                        workSheet = excel.Application.Worksheets.Add();
                        workSheet.Move(After: excel.Sheets[excel.Sheets.Count]);
                        workSheet.Name = sheetName;

                        // Add time range to excel sheet
                        TimeSpan timeStamp = new TimeSpan(0, 0, 0);
                        var rowLength = results.GetLength(1);
                        for (int k = 1; k <= rowLength; k++)
                        {
                            if (k == 1)
                            {
                                workSheet.Cells[k, 1] = "Time";
                            }
                            else
                            {
                                workSheet.Cells[k, 1] = timeStamp.ToString();
                                timeStamp = timeStamp.Add(TimeSpan.FromSeconds(time));
                            }
                        }
                        colomn = 2;
                        
                    }
                    var row = j + 1;

                    // Add results to excell sheet
                    if (j == 0)
                    {
                        workSheet.Cells[row, colomn] = "Results";
                    }
                    else if (results[i, j] != null)
                    {
                        workSheet.Cells[row, colomn] = results[i, j];
                    }
                }
            }

            var error = new ErrorHandler();

            try
            {
                // Save excel sheet
                workSheet.SaveAs(reportFile);
            }
            catch (Exception)
            {
                excel.Quit();
                error.Exit(93);
            }

            // Close sheet
            excel.Quit();
        }

        public void AddCharts(string reportFile)
        {
            var excel = new Excel.Application() {DisplayAlerts = false};
            excel.Workbooks.Open(reportFile);
            Excel._Worksheet workSheet = (Excel.Worksheet)excel.ActiveSheet;

            var sheetNumber = 1;
            var chartPositionNumber = 2;
            foreach (Excel.Worksheet sheet in excel.Worksheets)
            {
                if (sheet.Name != "Charts")
                {
                    // Set the datasheet for the source of the data
                    Excel.Worksheet dataSheet = excel.Worksheets[sheetNumber];

                    // Open the chart sheet to save the charts
                    workSheet = excel.ActiveWorkbook.Sheets["Charts"];
                    workSheet.Select();

                    // Get column & row length
                    var colomn = dataSheet.UsedRange.Columns.Count;
                    var rows = dataSheet.UsedRange.Rows.Count - 1;

                    // Chart settings and stuff
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
                    Excel.ChartObject runChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);

                    Excel.Chart runChartPage = runChart.Chart;

                    runChartPage.ChartType = Excel.XlChartType.xlLine;

                    // set ChartStyle based on Office version
                    var chartStyle = 301;
                    if (Type.GetTypeFromProgID("Excel.Application.14") != null)
                    {
                        chartStyle = 2;
                    }

                    runChartPage.HasTitle = true;
                    runChartPage.HasLegend = true;
                    runChartPage.ChartTitle.Text = sheet.Name;
                    runChartPage.ChartStyle = chartStyle;

                    // Position of chart
                    var runChartPosition = "B" + chartPositionNumber;
                    Excel.Range runChartPlacementRange = workSheet.get_Range(runChartPosition, runChartPosition);

                    runChart.Top = runChartPlacementRange.Top;
                    runChart.Left = runChartPlacementRange.Left;

                    chartPositionNumber = chartPositionNumber + 21;

                    // Size of Chart
                    runChart.Width = 500;
                    runChart.Height = 250;
                    Excel.SeriesCollection runSeriesCollection = runChartPage.SeriesCollection();

                    // Create run line chart
                    for (int i = 2; i <= (colomn); i++)
                    {

                        Excel.Series runSeries = runSeriesCollection.NewSeries();
                        runSeries.Name = dataSheet.Cells[1, i].Value;


                        // set correct range for chart data
                        var ia = i;
                        // Time range
                        var xValuesBegin = ParseColumnName(1) + "2";
                        var xValuesEnd = ParseColumnName(1) + (rows.ToString());

                        var valuesBegin = ParseColumnName(ia) + (2).ToString();
                        var valuesEnd = ParseColumnName(ia) + (rows + 1).ToString();

                        runSeries.XValues = dataSheet.get_Range(xValuesBegin, xValuesEnd);
                        runSeries.Values = dataSheet.get_Range(valuesBegin, valuesEnd);
                    }
                }
                sheetNumber++;
            }

            var error = new ErrorHandler();

            try
            {
                // Save excel sheet
                workSheet.SaveAs(reportFile);
            }
            catch (Exception)
            {
                excel.Quit();
                error.Exit(93);
            }

            // Close sheet
            excel.Quit();
        }

        public string ParseColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
