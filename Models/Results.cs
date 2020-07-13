using System.Collections.Generic;

namespace Parse_Performance_Data.Models
{
    public class Results
    {
        public string File { get; set; }
        public string Header { get; set; }
        public List<string> Data { get; set; }
    }
}
