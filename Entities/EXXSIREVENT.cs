using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("EXX_SIRE_VENT")]
    public class EXXSIREVENT
    {
        [JsonProperty("DocEntry"), SAPKey] public string DocEntry { get; set; }
        [JsonProperty("U_EXX_PERIODO")] public string UEXXPERIODO { get; set; } //Periodo
        [JsonProperty("U_EXX_FECREG")] public string UEXXFECREG { get; set; } //Apellidos y Nombres o Razón social
        [JsonProperty("U_EXX_ARCHTXT")] public string UEXXARCHTXT { get; set; } //Periodo
        [JsonProperty("U_EXX_SIRE_VENT1Collection")] public EXXSIREVENT1 EXXSIREVENT1 { get; set; } //
    }

    public class EXXSIREVENT1
    {
        [JsonProperty("U_EXX_ORIGEN")] public int UEXXORIGEN { get; set; } // 1 SAP Y SIRE  - 2 SAP - 3 SIRE
        [JsonProperty("U_EXX_OBJTYPE")] public string UEXXOBJTYPE { get; set; } //Tipo SAP
        [JsonProperty("U_EXX_DOCENTRY")] public string UEXXDOCENTRY { get; set; } //Número SAP
        [JsonProperty("U_EXX_RUC")] public string UEXXRUC { get; set; } //RUC
        [JsonProperty("U_EXX_RAZSOC")] public string UEXXRAZSOC { get; set; } //Apellidos y Nombres o Razon social
        [JsonProperty("U_EXX_PERIODO")] public string UEXXPERIODO { get; set; } //Periodo
        [JsonProperty("U_EXX_CARSUN")] public string UEXXCARSUN { get; set; } //CAR SUNAT
        [JsonProperty("U_EXX_FEMI")] public string UEXXFEMI { get; set; } //Fecha de emision
        [JsonProperty("U_EXX_FVCTO")] public string UEXXFVCTO { get; set; } //Fecha Vcto/Pago
        [JsonProperty("U_EXX_TIPDOC")] public string UEXXTIPDOC { get; set; } //Tipo CP/Doc
        [JsonProperty("U_EXX_SERIE")] public string UEXXSERIE { get; set; } //Serie del CP
        [JsonProperty("U_EXX_NROINI")] public string UEXXNROINI { get; set; } //Nro Inicial Rango
        [JsonProperty("U_EXX_NROFIN")] public string UEXXNROFIN { get; set; } //Nro Final Rango
        [JsonProperty("U_EXX_TIPIDE")] public string UEXXTIPIDE { get; set; } //Tipo Doc Identidad
        [JsonProperty("U_EXX_NUMIDE")] public string UEXXNUMIDE { get; set; } //Nro Doc Identidad
        [JsonProperty("U_EXX_RAZCLI")] public string UEXXRAZCLI { get; set; } //Apellidos Nombres o Razon Social
        [JsonProperty("U_EXX_VFAEX")] public double? UEXXVFAEX { get; set; } //Valor Facturado Exportación
        [JsonProperty("U_EXX_BIGRA")] public double? UEXXBIGRA { get; set; } //BI Gravado
        [JsonProperty("U_EXX_DESBI")] public double? UEXXDESBI { get; set; } //Dscto BI
        [JsonProperty("U_EXX_IGVIPM")] public double? UEXXIGVIPM { get; set; } //IGV/IPM
        [JsonProperty("U_EXX_DESII")] public double? UEXXDESII { get; set; } //Dscto IGV/IPM
        [JsonProperty("U_EXX_MOEX")] public double? UEXXMOEX { get; set; } //Monto Exonerado
        [JsonProperty("U_EXX_MOIN")] public double? UEXXMOIN { get; set; } //Monto Inafecto
        [JsonProperty("U_EXX_ISC")] public double? UEXXISC { get; set; } //ISC
        [JsonProperty("U_EXX_BIGIP")] public double? UEXXBIGIP { get; set; } //BI Grav IVAP
        [JsonProperty("U_EXX_IVAP")] public double? UEXXIVAP { get; set; } //IVAP
        [JsonProperty("U_EXX_ICBPER")] public double? UEXXICBPER { get; set; } //ICBPER
        [JsonProperty("U_EXX_OTRIB")] public double? UEXXOTRIB { get; set; } //Otros Tributos
        [JsonProperty("U_EXX_TOTAL")] public double? UEXXTOTAL { get; set; } //Total CP

        [JsonProperty("U_EXX_MONEDA")] public string UEXXMONEDA { get; set; } //Moneda
        [JsonProperty("U_EXX_TCAMBIO")] public string UEXXTCAMBIO { get; set; } //Tipo de Cambio

        [JsonProperty("U_EXX_FEMOD")] public string UEXXFEMOD { get; set; } //Fecha Emision Doc Modificado
        [JsonProperty("U_EXX_TIPMOD")] public string UEXXTIPMOD { get; set; } //Tipo CP Modificado
        [JsonProperty("U_EXX_SERMOD")] public string UEXXSERMOD { get; set; } //Serie CP Modificado
        [JsonProperty("U_EXX_NUMMOD")] public string UEXXNUMMOD { get; set; } //Nro CP Modificado

        [JsonProperty("U_EXX_IDPOA")] public string UEXXIDPOA { get; set; } //ID Proy. Operadores Atrib.
        [JsonProperty("U_EXX_TIPNOT")] public string UEXXTIPNOT { get; set; } //Tipo de Nota
        [JsonProperty("U_EXX_ESTCOM")] public string UEXXESTCOM { get; set; } //Est. Comp

        [JsonProperty("U_EXX_VFOBE")] public double? UEXXVFOBE { get; set; } //Valor FOB Embarcado
        [JsonProperty("U_EXX_VOPGRA")] public double? UEXXVOPGRA { get; set; } //Valor OP Gratuitas

        [JsonProperty("U_EXX_TIPOPE")] public string UEXXTIPOPE { get; set; } //Tipo Operación
        [JsonProperty("U_EXX_DAMCP")] public string UEXXDAMCP { get; set; } //DAM/CP
        [JsonProperty("U_EXX_CLU1")] public string UEXXCLU1 { get; set; } //CLU1

    }
}
