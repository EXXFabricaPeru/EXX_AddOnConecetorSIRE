using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    public class Ticket
    {
        [JsonProperty("paginacion")] public Paginacion Paginacion { get; set; }
        [JsonProperty("registros")] public List<Registro> Registros { get; set; }
    }

    public class Paginacion
    {
        [JsonProperty("page")] public int Page { get; set; }
        [JsonProperty("perPage")] public int PerPage { get; set; }
        [JsonProperty("totalRegistros")] public int TotalRegistros { get; set; }
    }

    public class Registro
    {
        [JsonProperty("showReporteDescarga")] public string ShowReporteDescarga { get; set; }
        [JsonProperty("perTributario")] public string PerTributario { get; set; }
        [JsonProperty("numTicket")] public string NumTicket { get; set; }
        [JsonProperty("fecCargaImportacion")] public DateTime? FecCargaImportacion { get; set; }
        [JsonProperty("fecInicioProceso")] public DateTime? FecInicioProceso { get; set; }
        [JsonProperty("codProceso")] public string CodProceso { get; set; }
        [JsonProperty("desProceso")] public string DesProceso { get; set; }
        [JsonProperty("codEstadoProceso")] public string CodEstadoProceso { get; set; }
        [JsonProperty("desEstadoProceso")] public string DesEstadoProceso { get; set; }
        [JsonProperty("nomArchivoImportacion")] public string NomArchivoImportacion { get; set; }
        [JsonProperty("detalleTicket")] public DetalleTicket DetalleTicket { get; set; }
        [JsonProperty("archivoReporte")] public List<ArchivoReporte> ArchivoReporte { get; set; }
        [JsonProperty("subProcesos")] public List<SubProceso> SubProcesos { get; set; }
    }

    public class DetalleTicket
    {
        [JsonProperty("numTicket")] public string NumTicket { get; set; }
        [JsonProperty("fecCargaImportacion")] public DateTime? FecCargaImportacion { get; set; }
        [JsonProperty("horaCargaImportacion")] public string HoraCargaImportacion { get; set; }
        [JsonProperty("codEstadoEnvio")] public string CodEstadoEnvio { get; set; }
        [JsonProperty("desEstadoEnvio")] public string DesEstadoEnvio { get; set; }
        [JsonProperty("nomArchivoReporte")] public string NomArchivoReporte { get; set; }
        [JsonProperty("cntFilasvalidada")] public int CntFilasValidada { get; set; }
        [JsonProperty("cntCPError")] public int CntCPError { get; set; }
        [JsonProperty("cntCPInformados")] public int CntCPInformados { get; set; }
    }

    public class ArchivoReporte
    {
        [JsonProperty("codTipoAchivoReporte")] public string CodTipoAchivoReporte { get; set; }
        [JsonProperty("nomArchivoReporte")] public string NomArchivoReporte { get; set; }
        [JsonProperty("nomArchivoContenido")] public string NomArchivoContenido { get; set; }
    }

    public class SubProceso
    {
        [JsonProperty("codTipoSubProceso")] public string CodTipoSubProceso { get; set; }
        [JsonProperty("desTipoSubProceso")] public string DesTipoSubProceso { get; set; }
        [JsonProperty("codEstado")] public string CodEstado { get; set; }
        [JsonProperty("numIntentos")] public int NumIntentos { get; set; }
    }
}
