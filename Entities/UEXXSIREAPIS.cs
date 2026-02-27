using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("U_EXX_SIRE_APIS")]
    public class UEXXSIREAPIS
    {
        [JsonProperty("Code"), SAPKey] public string Code { get; set; }
        [JsonProperty("U_EXX_RUC")] public string UEXXRUC { get; set; } //RUC EMPRESA
        [JsonProperty("U_EXX_APIS")] public string UEXXAPIS { get; set; } //URL API SIRE
        [JsonProperty("U_EXX_USER")] public string UEXXUSER { get; set; } //Usuario SOL
        [JsonProperty("U_EXX_PASS")] public string UEXXPASS { get; set; } //Clave SAP
        [JsonProperty("U_EXX_CLID")] public string UEXXCLID { get; set; } //Client id
        [JsonProperty("U_EXX_CLSE")] public string UEXXCLSE { get; set; } //Client secret
    }
}
