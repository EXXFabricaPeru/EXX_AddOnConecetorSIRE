using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("UserObjectsMD")]
    public class UserObjectsMd
    {
        [JsonProperty("Code"), SAPKey] public string Code { get; set; }
        [JsonProperty("TableName")] public string TableName { get; set; }
        [JsonProperty("LogTableName")] public string LogTableName { get; set; }
        [JsonProperty("CanCreateDefaultForm")] public string CanCreateDefaultForm { get; set; }
        [JsonProperty("ObjectType")] public string ObjectType { get; set; }
        [JsonProperty("CanCancel")] public string CanCancel { get; set; }
        [JsonProperty("CanDelete")] public string CanDelete { get; set; }
        [JsonProperty("CanLog")] public string CanLog { get; set; }
        [JsonProperty("ManageSeries")] public string ManageSeries { get; set; }
        [JsonProperty("CanFind")] public string CanFind { get; set; }
        [JsonProperty("CanYearTransfer")] public string CanYearTransfer { get; set; }
        [JsonProperty("Name")] public string Name { get; set; }
        [JsonProperty("CanClose")] public string CanClose { get; set; }
        [JsonProperty("CanArchive")] public string CanArchive { get; set; }
        [JsonProperty("EnableEnhancedForm")] public string EnableEnhancedForm { get; set; }
        [JsonProperty("RebuildEnhancedForm")] public string RebuildEnhancedForm { get; set; }
        [JsonProperty("FormSRF")] public string FormSRF { get; set; }
        [JsonProperty("ApplyAuthorization")] public string ApplyAuthorization { get; set; }
        [JsonProperty("UserObjectMD_ChildTables")] public List<UserObjectMDChildTable> UserObjectMD_ChildTables { get; set; }
        [JsonProperty("UserObjectMD_FindColumns")] public List<UserObjectMDFindColumn> UserObjectMD_FindColumns { get; set; }
        [JsonProperty("UserObjectMD_FormColumns")] public List<UserObjectMDFormColumn> UserObjectMD_FormColumns { get; set; }
    }

    public class UserObjectMDChildTable
    {
        [JsonProperty("SonNumber")] public int? SonNumber { get; set; }
        [JsonProperty("TableName")] public string TableName { get; set; }
        [JsonProperty("LogTableName")] public string LogTableName { get; set; }
        [JsonProperty("Code")] public string Code { get; set; }
        [JsonProperty("ObjectName")] public string ObjectName { get; set; }
    }

    public class UserObjectMDFindColumn
    {
        [JsonProperty("ColumnNumber")] public int? ColumnNumber { get; set; }
        [JsonProperty("ColumnAlias")] public string ColumnAlias { get; set; }
        [JsonProperty("ColumnDescription")] public string ColumnDescription { get; set; }
        [JsonProperty("Code")] public string Code { get; set; }
    }

    public class UserObjectMDFormColumn
    {
        [JsonProperty("FormColumnAlias")] public string FormColumnAlias { get; set; }
        [JsonProperty("FormColumnDescription")] public string FormColumnDescription { get; set; }
        [JsonProperty("FormColumnNumber")] public int? FormColumnNumber { get; set; }
        [JsonProperty("SonNumber")] public int? SonNumber { get; set; }
        [JsonProperty("Code")] public string Code { get; set; }
        [JsonProperty("Editable")] public string Editable { get; set; }
    }
}
