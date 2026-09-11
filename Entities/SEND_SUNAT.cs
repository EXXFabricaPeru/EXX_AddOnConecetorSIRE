using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    public class SEND_SUNAT
    {
        [JsonProperty("numRuc")] public string numRuc { get; set; }
        [JsonProperty("codComp")] public string codComp { get; set; }
        [JsonProperty("numeroSerie")] public string numeroSerie { get; set; }
        [JsonProperty("numero")] public int numero { get; set; }
        [JsonProperty("fechaEmision")] public string fechaEmision { get; set; }
        [JsonProperty("monto")] public double monto { get; set; }
    }
}
