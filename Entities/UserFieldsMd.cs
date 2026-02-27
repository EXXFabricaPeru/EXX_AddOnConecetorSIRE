using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("UserFieldsMD")]
    public class UserFieldsMd
    {
        [JsonProperty("TableName")] public string TableName { get; set; }
        [JsonProperty("Name"), SAPKey] public string Name { get; set; }
        [JsonProperty("Description")] public string Description { get; set; }
        [JsonProperty("Type")] public string Type { get; set; }
        [JsonProperty("SubType")] public string SubType { get; set; }
        [JsonProperty("EditSize")] public int? EditSize { get; set; }
        [JsonProperty("LinkedSystemObject")] public string LinkedSystemObject { get; set; }
        [JsonProperty("DefaultValue")] public string DefaultValue { get; set; }
        [JsonProperty("LinkedUDO")] public string LinkedUDO { get; set; }
        [JsonProperty("LinkedTable")] public string LinkedTable { get; set; }
        [JsonProperty("Mandatory")] public string Mandatory { get; set; }
        [JsonProperty("ValidValuesMD")] public List<ValidValuesMd> ValidValuesMD { get; set; }
    }

    public class ValidValuesMd
    {
        [JsonProperty("Value"), SAPKey] public string Value { get; set; }
        [JsonProperty("Description")] public string Description { get; set; }
    }
}
