using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("UserTablesMD")]
    public class UserTablesMd
    {
        [JsonProperty("TableName"), SAPKey] public string TableName { get; set; }
        [JsonProperty("TableDescription")] public string TableDescription { get; set; }
        [JsonProperty("TableType")] public string TableType { get; set; }
        [JsonProperty("Archivable")] public string Archivable { get; set; }
        [JsonProperty("ArchiveDateField")] public string ArchiveDateField { get; set; }
    }
}
