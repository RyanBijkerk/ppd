using System;

namespace Parse_Performance_Data
{
    public class ErrorHandler
    {
        public void Exit(int code)
        {
            switch (code)
            {
                case 2:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: The system cannot find the file specified.", DateTime.Now,
                        code);
                    break;
                case 3:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: The system cannot find the location specified.");
                    break;
                case 4:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: The system cannot open the file.");
                    break;
                case 13:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: The data is invalid.");
                    break;
                case 93:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: File in use.");
                    break;
                case 160:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: One or more arguments are not correct.");
                    break;
                case 161:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: Invalid precentage parameter.");
                    break;
                case 162:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: Precentage parameter must be between 0 and 100.");
                    break;
                case 183:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: Cannot create a file when that file already exists.");
                    break;
                default:
                    Console.WriteLine($"{DateTime.Now}: Error code {code}:: Unknown error code.");
                    break;
            }

            Environment.Exit(code);
        }

    }
}
