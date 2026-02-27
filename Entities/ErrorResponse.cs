using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    public class ErrorResponse
    {
        public error Error { get; set; }
    }

    public class error
    {
        public string code { get; set; }
        public message Message { get; set; }
    }

    public class message
    {
        public string lang { get; set; }
        public string value { get; set; }
    }
}
