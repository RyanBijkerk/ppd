using System.Collections.Generic;
using Parse_Performance_Data.Modal;

namespace Parse_Performance_Data.Code
{
    class ArgumentHandler
    {
        public Arguments Parse(string[] args)
        {
            var arguments = new Arguments();
            var error = new ErrorHandler();
            var help = new HelpHandler();

            // set default merge item
            arguments.Merge = false;

            // Check for arguments defined
            if (args.GetLength(0) == 0)
            {
                error.Exit(160);
            }
            var collectedFile = new List<string>();
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "/?":
                        help.Instuctions();
                        break;
                    case "help":
                        help.Instuctions();
                        break;
                    case "?":
                        help.Instuctions();
                        break;
                    case "/f":
                        collectedFile.Add(args[i + 1]);
                        break;
                    case "/merge":
                        arguments.Merge = true;
                        break;
                }
            }
            // Add the list to the modal
            arguments.fileLocation = collectedFile;

            // See if the filelocation argument is defined
            if (arguments.fileLocation == null)
            {
                error.Exit(160);
            }

            return arguments;
        }
    }
}
