using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Entities
{
    [SAPEntityName("EXX_SIRE_COMP")]
    public class EXXSIRECOMP
    {
        [JsonProperty("DocEntry"), SAPKey] public string DocEntry { get; set; }
        [JsonProperty("U_EXX_PERIODO")] public string UEXXPERIODO { get; set; } //RUC
        [JsonProperty("U_EXX_FECREG")] public string UEXXFECREG { get; set; } //Apellidos y Nombres o Razón social
        [JsonProperty("U_EXX_ARCHTXT")] public string UEXXARCHTXT { get; set; } //Periodo
        [JsonProperty("U_EXX_SIRE_COMP1Collection")] public EXXSIRECOMP1 EXXSIRECOMP1 { get; set; } //Periodo
    }

    public class EXXSIRECOMP1
    {
        [JsonProperty("U_EXX_ORIGEN")] public int UEXXORIGEN { get; set; } // 1 SAP Y SIRE  - 2 SAP - 3 SIRE
        [JsonProperty("U_EXX_OBJTYPE")] public string UEXXOBJTYPE { get; set; }
        [JsonProperty("U_EXX_DOCENTRY")] public string UEXXDOCENTRY { get; set; }
        [JsonProperty("U_EXX_RUC")] public string UEXXRUC { get; set; } //RUC
        [JsonProperty("U_EXX_RAZSOC")] public string UEXXRAZSOC { get; set; } //Apellidos y Nombres o Razón social
        [JsonProperty("U_EXX_PERIODO")] public string UEXXPERIODO { get; set; } //Periodo
        [JsonProperty("U_EXX_CARSUN")] public string UEXXCARSUN { get; set; } //CAR SUNAT
        [JsonProperty("U_EXX_FEMI")] public string UEXXFEMI { get; set; } //Fecha de emisión
        [JsonProperty("U_EXX_FVCTO")] public string UEXXFVCTO { get; set; } //Fecha Vcto/Pago
        [JsonProperty("U_EXX_TIPDOC")] public string UEXXTIPDOC { get; set; } //Tipo CP/Doc.
        [JsonProperty("U_EXX_SERIE")] public string UEXXSERIE { get; set; } //Serie del CDP
        [JsonProperty("U_EXX_ANIO")] public string UEXXANIO { get; set; } //Año
        [JsonProperty("U_EXX_NROINI")] public string UEXXNROINI { get; set; } //Nro Inicial (Rango)
        [JsonProperty("U_EXX_NROFIN")] public string UEXXNROFIN { get; set; } //Nro Final (Rango)
        [JsonProperty("U_EXX_TIPIDE")] public string UEXXTIPIDE { get; set; } //Tipo Doc Identidad
        [JsonProperty("U_EXX_NUMIDE")] public string UEXXNUMIDE { get; set; } //Nro Doc Identidad
        [JsonProperty("U_EXX_RAZCLI")] public string UEXXRAZCLI { get; set; } //Apellidos Nombres/ Razón  Social
        [JsonProperty("U_EXX_BIGDG")] public double? UEXXBIGDG { get; set; } //BI Gravado DG
        [JsonProperty("U_EXX_IGVDG")] public double? UEXXIGVDG { get; set; } //IGV / IPM DG
        [JsonProperty("U_EXX_BIGDGN")] public double? UEXXBIGDGN { get; set; } //BI Gravado DGNG
        [JsonProperty("U_EXX_IGVDGN")] public double? UEXXIGVDGN { get; set; } //IGV / IPM DGNG
        [JsonProperty("U_EXX_BIDNG")] public double? UEXXBIDNG { get; set; } //BI Gravado DNG
        [JsonProperty("U_EXX_IGVDNG")] public double? UEXXIGVDNG { get; set; } //IGV / IPM DNG
        [JsonProperty("U_EXX_VALNG")] public double? UEXXVALNG { get; set; } //Valor Adq. NG
        [JsonProperty("U_EXX_ISC")] public double? UEXXISC { get; set; } //ISC
        [JsonProperty("U_EXX_ICBPER")] public double? UEXXICBPER { get; set; } //ICBPER
        [JsonProperty("U_EXX_OTRTRI")] public double? UEXXOTRTRI { get; set; } //Otros Trib/ Cargos
        [JsonProperty("U_EXX_TOTAL")] public double? UEXXTOTAL { get; set; } //Total CP
        [JsonProperty("U_EXX_MONEDA")] public string UEXXMONEDA { get; set; } //Moneda
        [JsonProperty("U_EXX_TCAMBIO")] public string UEXXTCAMBIO { get; set; } //Tipo de Cambio
        [JsonProperty("U_EXX_FEMOD")] public string UEXXFEMOD { get; set; } //Fecha Emisión Doc Modificado
        [JsonProperty("U_EXX_TIPMOD")] public string UEXXTIPMOD { get; set; } //Tipo CP Modificado
        [JsonProperty("U_EXX_SERMOD")] public string UEXXSERMOD { get; set; } //Serie CP Modificado
        [JsonProperty("U_EXX_CODDAM")] public string UEXXCODDAM { get; set; } //COD. DAM O DSI
        [JsonProperty("U_EXX_NUMMOD")] public string UEXXNUMMOD { get; set; } //Nro CP Modificado
        [JsonProperty("U_EXX_CLASIF")] public string UEXXCLASIF { get; set; } //Clasif de Bss y Sss
        [JsonProperty("U_EXX_IDPROY")] public string UEXXIDPROY { get; set; } //ID Proyecto Operadores
        [JsonProperty("U_EXX_PORPAR")] public double? UEXXPORPAR { get; set; } //PorcPart
        [JsonProperty("U_EXX_IMB")] public double? UEXXIMB { get; set; } //IMB
        [JsonProperty("U_EXX_CARORI")] public string UEXXCARORI { get; set; } //CAR Orig/ Ind E o I
        [JsonProperty("U_EXX_DETRA")] public double? UEXXDETRA { get; set; } //Detracción
        [JsonProperty("U_EXX_TIPNOT")] public string UEXXTIPNOT { get; set; } //Tipo de Nota
        [JsonProperty("U_EXX_ESTCOM")] public string UEXXESTCOM { get; set; } //Est. Comp.
        [JsonProperty("U_EXX_INCAL")] public string UEXXINCAL { get; set; } //Incal
        [JsonProperty("U_EXX_CLU1")] public string UEXXCLU1 { get; set; } //CLU1
        [JsonProperty("U_EXX_CLU2")] public string UEXXCLU2 { get; set; } //CLU2
        [JsonProperty("U_EXX_CLU3")] public string UEXXCLU3 { get; set; } //CLU3
        [JsonProperty("U_EXX_CLU4")] public string UEXXCLU4 { get; set; } //CLU4
        [JsonProperty("U_EXX_CLU5")] public string UEXXCLU5 { get; set; } //CLU5
        [JsonProperty("U_EXX_CLU6")] public string UEXXCLU6 { get; set; } //CLU6
        [JsonProperty("U_EXX_CLU7")] public string UEXXCLU7 { get; set; } //CLU7
        [JsonProperty("U_EXX_CLU8")] public string UEXXCLU8 { get; set; } //CLU8
        [JsonProperty("U_EXX_CLU9")] public string UEXXCLU9 { get; set; } //CLU9
        [JsonProperty("U_EXX_CLU10")] public string UEXXCLU10 { get; set; } //CLU10
        [JsonProperty("U_EXX_CLU11")] public string UEXXCLU11 { get; set; } //CLU11
        [JsonProperty("U_EXX_CLU12")] public string UEXXCLU12 { get; set; } //CLU12
        [JsonProperty("U_EXX_CLU13")] public string UEXXCLU13 { get; set; } //CLU13
        [JsonProperty("U_EXX_CLU14")] public string UEXXCLU14 { get; set; } //CLU14
        [JsonProperty("U_EXX_CLU15")] public string UEXXCLU15 { get; set; } //CLU15
        [JsonProperty("U_EXX_CLU16")] public string UEXXCLU16 { get; set; } //CLU16
        [JsonProperty("U_EXX_CLU17")] public string UEXXCLU17 { get; set; } //CLU17
        [JsonProperty("U_EXX_CLU18")] public string UEXXCLU18 { get; set; } //CLU18
        [JsonProperty("U_EXX_CLU19")] public string UEXXCLU19 { get; set; } //CLU19
        [JsonProperty("U_EXX_CLU20")] public string UEXXCLU20 { get; set; } //CLU20
        [JsonProperty("U_EXX_CLU21")] public string UEXXCLU21 { get; set; } //CLU21
        [JsonProperty("U_EXX_CLU22")] public string UEXXCLU22 { get; set; } //CLU22
        [JsonProperty("U_EXX_CLU23")] public string UEXXCLU23 { get; set; } //CLU23
        [JsonProperty("U_EXX_CLU24")] public string UEXXCLU24 { get; set; } //CLU24
        [JsonProperty("U_EXX_CLU25")] public string UEXXCLU25 { get; set; } //CLU25
        [JsonProperty("U_EXX_CLU26")] public string UEXXCLU26 { get; set; } //CLU26
        [JsonProperty("U_EXX_CLU27")] public string UEXXCLU27 { get; set; } //CLU27
        [JsonProperty("U_EXX_CLU28")] public string UEXXCLU28 { get; set; } //CLU28
        [JsonProperty("U_EXX_CLU29")] public string UEXXCLU29 { get; set; } //CLU29
        [JsonProperty("U_EXX_CLU30")] public string UEXXCLU30 { get; set; } //CLU30
        [JsonProperty("U_EXX_CLU31")] public string UEXXCLU31 { get; set; } //CLU31
        [JsonProperty("U_EXX_CLU32")] public string UEXXCLU32 { get; set; } //CLU32
        [JsonProperty("U_EXX_CLU33")] public string UEXXCLU33 { get; set; } //CLU33
        [JsonProperty("U_EXX_CLU34")] public string UEXXCLU34 { get; set; } //CLU34
        [JsonProperty("U_EXX_CLU35")] public string UEXXCLU35 { get; set; } //CLU35
        [JsonProperty("U_EXX_CLU36")] public string UEXXCLU36 { get; set; } //CLU36
        [JsonProperty("U_EXX_CLU37")] public string UEXXCLU37 { get; set; } //CLU37
        [JsonProperty("U_EXX_CLU38")] public string UEXXCLU38 { get; set; } //CLU38
        [JsonProperty("U_EXX_CLU39")] public string UEXXCLU39 { get; set; } //CLU39
    }
}
