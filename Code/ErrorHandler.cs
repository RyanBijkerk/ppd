using System;

namespace Parse_Performance_Data
{
    class ErrorHandler
    {
        public void Exit(int Code)
        {
            if (Code == 2)
            {
                Console.WriteLine("{0}: Error code {1}: The system cannot find the file specified.", DateTime.Now, Code);
                Environment.Exit(Code);
            }
            
            if (Code == 4)
            {
                Console.WriteLine("{0}: Error code {1}: The system cannot open the file.", DateTime.Now, Code);
                Environment.Exit(Code);
            }

            if (Code == 13)
            {
                Console.WriteLine("{0}: Error code {1}: The data is invalid.", DateTime.Now, Code);
                Environment.Exit(Code);
            }

            if (Code == 93)
            {
                Console.WriteLine("{0}: Error code {1}: File in use.", DateTime.Now, Code);
                Environment.Exit(Code);
            }


            if (Code == 160)
            {
                Console.WriteLine("{0}: Error code {1}: One or more arguments are not correct.", DateTime.Now, Code);
                Environment.Exit(Code);
            }

            if (Code == 183)
            {
                Console.WriteLine("{0}: Error code {1}: Cannot create a file when that file already exists.", DateTime.Now, Code);
                Environment.Exit(Code);
            }

            if (Code == 2013)
            {
                Console.WriteLine("{0}: Error code {1}: Microsoft Excel is missing.", DateTime.Now, Code);
                Environment.Exit(Code);
            }

        }

    }
}
