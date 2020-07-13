using System;

namespace Parse_Performance_Data.Code
{
    class HelpHandler
    {
        public void Instuctions()
        {
            Console.WriteLine("Usage: PPD.exe /f @file");
            Console.WriteLine("Example: PPD.exe /f C:\\Temp\\Performance.csv");
            Console.WriteLine("");
            Console.WriteLine("Requires: Microsoft Excel 2010 or higher");
            Environment.Exit(0);            
        }
    }
}
