using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("U_EXX_SETUP")]
    public class UEXXSETUP
    {
        [JsonProperty("Code"), SAPKey] public string Code { get; set; }
        [JsonProperty("Name")] public string Name { get; set; }
        [JsonProperty("U_EXX_ADDN")] public string U_EXX_ADDN { get; set; }
        [JsonProperty("U_EXX_VERS")] public string U_EXX_VERS { get; set; }
        [JsonProperty("U_EXX_RUTA")] public string U_EXX_RUTA { get; set; }
    }
}
