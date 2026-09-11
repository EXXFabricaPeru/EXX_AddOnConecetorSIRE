using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    public class Propuesta
    {
        [JsonProperty("numTicket")] public string numTicket { get; set; }
    }
}
