using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("U_EXX_SIRE_CONF")]
    public class UEXXSIRECONF
    {
        [JsonProperty("Code"), SAPKey] public string Code { get; set; }
        [JsonProperty("U_EXX_CONS")] public string UEXXCONS { get; set; } //Tipo Conecxion SAP
        [JsonProperty("U_EXX_URSL")] public string UEXXURSL { get; set; } //URL Service Layer
        [JsonProperty("U_EXX_VSAP")] public string UEXXVSAP { get; set; } //Valida compras SAP: 1-TXT | 2-Store procedure
        [JsonProperty("U_EXX_VSIR")] public string UEXXVSIR { get; set; } //Valida compras SAP: 1-TXT | 2-API SIRE
        public List<UEXXSIREAPIS> APIS { get; set; } //Valida compras SAP: 1-TXT | 2-API SIRE
    }
}
