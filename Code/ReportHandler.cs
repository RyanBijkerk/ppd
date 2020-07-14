using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using Parse_Performance_Data.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Parse_Performance_Data
{
    public class ReportHandler
    {
        private readonly ErrorHandler _errorHandler;

        public ReportHandler(ErrorHandler errorHandler)
        {
            _errorHandler = errorHandler;
        }

        public ReportHandler()
        {
        }

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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var cleanFile = Path.GetFileNameWithoutExtension(file);
            var location = Path.GetDirectoryName(file);
            var reportFile = location + "\\" + cleanFile + ".xlsx";

            if (!File.Exists(reportFile))
            {
                using (var fs = new FileStream(reportFile, FileMode.Create, FileAccess.Write))
                {
                    var excel = new ExcelPackage();

                    excel.Workbook.Properties.Author = "Ryan Ververs-Bijkerk";
                    excel.Workbook.Properties.Company = "Logit Blog";
                    excel.Workbook.Worksheets.Add("Charts");

                    excel.SaveAs(fs);
                }
            }

            return reportFile;
        }

        public void AddData(string reportFile, List<Results> results, double time, int lines)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(reportFile);
            using (var excel = new ExcelPackage(file))
            {
                var workSheet = excel.Workbook.Worksheets.FirstOrDefault();
                var colomn = 0;

                foreach (var result in results)
                {
                    var sheetName = result.Header;

                    if (sheetName.Contains('\\'))
                    {
                        var regex = new Regex(@"[\\]{2}(?<host>.+)\\.+\\(?<metric>.+)");
                        var match = regex.Match(sheetName);

                        if (match.Success)
                        {
                            sheetName = match.Groups[1].Value + "-" + match.Groups[2].Value;
                        }

                        if (sheetName.Contains("(vmhba"))
                        {
                            sheetName = sheetName.Substring(sheetName.LastIndexOf('('));
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
                    var maxLength = 31;
                    if (sheetName.Length > maxLength)
                    {
                        //var maxBeginPoint = sheetName.Length - maxLength;
                        sheetName = sheetName.Substring(0, maxLength);
                    }

                    // Check if worksheet exists and if not add new sheet
                    var sheetFound = false;

                    // Loop trough sheets
                    foreach (var sheet in excel.Workbook.Worksheets)
                    {
                        if (sheet.Name == sheetName)
                        {
                            // If found break from the loop
                            sheetFound = true;
                            break;
                        }
                    }
                    // If not found add the new sheet else set as active
                    if (!sheetFound)
                    {
                        workSheet = excel.Workbook.Worksheets.Add(sheetName);
                        excel.Workbook.Worksheets.MoveToEnd(sheetName);

                    }
                    else
                    {
                        workSheet = excel.Workbook.Worksheets.FirstOrDefault(w => w.Name == sheetName);
                    }


                    // Add time range to excel sheet
                    TimeSpan timeStamp = new TimeSpan(0, 0, 0);
                    var rowLength = result.Data.Count + 1;

                    // Check if time is set in the sheet
                    var header = (workSheet.Cells[1, 1]).Value;
                    if ((string)header != $"Time")
                    {
                        for (int k = 1; k <= rowLength; k++)
                        {
                            if (k == 1)
                            {
                                workSheet.Cells[k, 1].Value = "Time";
                            }
                            else
                            {
                                workSheet.Cells[k, 1].Value = timeStamp.ToString();
                                timeStamp = timeStamp.Add(TimeSpan.FromSeconds(time));
                            }
                        }
                    }

                    colomn = workSheet.Dimension.Columns + 1;

                    var row = 1;

                    workSheet.Cells[row, colomn].Value = "Results";
                    row++;

                    foreach (var data in result.Data)
                    {

                        try
                        {
                            workSheet.Cells[row, colomn].Value = Convert.ToDouble(data);
                        }
                        catch
                        {
                            workSheet.Cells[row, colomn].Value = data;
                        }
                        row++;
                    }
                }

                //for (int i = 0; i <= arrayLength; i++)
                //{
                //    var r = 0;
                //    for (int j = 0; j <= arraySecondLength; j++)
                //    {
                //        if (j == 0 && results[i, j] != null)
                //        {
                //            // Set the sheetname based on metric from last /
                //            var sheetName = results[i, j];

                //        }
                //        var row = r + 1;

                //        // Add results to excell sheet
                //        if (j == 0)
                //        {
                //            workSheet.Cells[row, colomn].Value = run;
                //        }
                //        else if (results[i, j] != null)
                //        {
                //            if (lines > 1)
                //            {
                //                var averageList = new List<double>();

                //                for (int av = 0; av < lines; av++)
                //                {
                //                    var avi = j + av;
                //                    var max = results.GetLength(1) - 1;
                //                    if (avi <= max)
                //                    {

                //                        try
                //                        {
                //                            averageList.Add(Convert.ToDouble(results[i, avi]));
                //                        }
                //                        catch
                //                        {
                //                            // ignore this result 
                //                        }
                //                    }
                //                }
                //                try
                //                {
                //                    var averageResult = averageList.Average();
                //                    workSheet.Cells[row, colomn].Value = (double) averageResult;
                //                }
                //                catch
                //                {
                //                    // ignore results
                //                }


                //                j = j + lines - 1;

                //            }
                //            else
                //            {
                //                try
                //                {
                //                    workSheet.Cells[row, colomn].Value = Convert.ToDouble(results[i, j]);
                //                }
                //                catch
                //                {
                //                    workSheet.Cells[row, colomn].Value = results[i, j];
                //                }


                //            }
                //        }

                //        r++;
                //    }
                //}

                var error = new ErrorHandler();

                try
                {
                    excel.SaveAs(file);
                }
                catch (Exception)
                {
                    error.Exit(93);
                }
            }
            
        }
        
        public void AddCharts(string reportFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(reportFile);
            using (var excel = new ExcelPackage(file))
            {
                var chartPositionNumber = 1;

                var workSheet = excel.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Charts");

                foreach (var sheet in excel.Workbook.Worksheets)
                {
                   
                    if (sheet.Name != "Charts")
                    {
                        var resultChart = workSheet.Drawings.AddLineChart(Guid.NewGuid().ToString(), eLineChartType.Line);
                        resultChart.SetPosition(chartPositionNumber, 0, 1, 0);
                        resultChart.SetSize(1300, 440);
                        resultChart.Title.Text = "Result " + sheet.Name;
                        resultChart.Legend.Add();
                        resultChart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle7);

                        chartPositionNumber = chartPositionNumber + 24;

                        var runRows = sheet.Dimension.Rows - 1;
                        var runColomns = sheet.Dimension.Columns - 1;
                        var runTime = sheet.Cells[2, 1, runRows, 1];

                        var resultData = sheet.Cells[2, (sheet.Dimension.Columns), (sheet.Dimension.Rows + 1), (sheet.Dimension.Columns)];
                        var resultSerie = resultChart.Series.Add(resultData, runTime);
                        resultSerie.Header = "Result";
                    }

                }

                try
                {
                    excel.Save();
                }
                catch (Exception)
                {
                    _errorHandler.Exit(93);
                }
            }
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
