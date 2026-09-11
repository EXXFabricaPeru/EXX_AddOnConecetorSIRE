using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Threading;
using Newtonsoft.Json;
using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;

namespace AddOnConectorSIRE.Modules.Asistente
{
    public class Main
    {
        public static void LoadForm()
        {
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            try
            {
                oForm = Globals.SBO_Application.Forms.Item("EXX_SIRE_ASIS");
                Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto.");
            }
            catch
            {
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = (SAPbouiCOM.FormCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "EXX_SIRE_ASIS";
                fcp.UniqueID = "EXX_SIRE_ASIS";
                string FormName = "\\SetupFields\\Forms\\SRF_EXX_SIRE_ASIS.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);
                int centerX = (Globals.SBO_Application.Desktop.Width - oForm.Width) / 2;
                int centerY = ((Globals.SBO_Application.Desktop.Height - Convert.ToInt32(Globals.SBO_Application.Desktop.Height * 0.15)) - oForm.Height) / 2;
                oForm.Left = centerX;
                oForm.Top = centerY;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");

                if (Globals.IsHana()) Globals.Query = $"SELECT \"Code\", \"Name\" FROM \"OFPR\" WHERE \"Code\" < '{DateTime.Now.ToString("yyyy-MM")}' ORDER BY \"Code\" DESC";
                else Globals.Query = $"SELECT Code, Name FROM [OFPR] WHERE Code < '{DateTime.Now.ToString("yyyy-MM")}' ORDER BY Code DESC";
                Globals.LlenarCombo(oForm, "PERIODO", Globals.Query, false, "");

                //if (Globals.IsHana())
                //{
                //    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxSucursal;
                //    else Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxOADM;
                //}
                //else
                //{
                //    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxSucursal;
                //    else Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxOADM;
                //}

                //oDataTable.ExecuteQuery(Globals.Query);
                //oMatrix.Columns.Item("BPLId").DataBind.Bind("DT_0", "BPLId");
                //oMatrix.Columns.Item("GlblLocNum").DataBind.Bind("DT_0", "GlblLocNum");
                //oMatrix.Columns.Item("BPLName").DataBind.Bind("DT_0", "BPLName");
                //oMatrix.Columns.Item("U_EXX_APIS").DataBind.Bind("DT_0", "U_EXX_APIS");
                //oMatrix.Columns.Item("U_EXX_USER").DataBind.Bind("DT_0", "U_EXX_USER");
                //oMatrix.Columns.Item("U_EXX_PASS").DataBind.Bind("DT_0", "U_EXX_PASS");
                //oMatrix.Columns.Item("U_EXX_CLID").DataBind.Bind("DT_0", "U_EXX_CLID");
                //oMatrix.Columns.Item("BPLId").DataBind.Bind("DT_0", "BPLId");
                //oMatrix.Columns.Item("U_EXX_CLSE").DataBind.Bind("DT_0", "U_EXX_CLSE");
                //oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
            }

            oForm.Refresh();
            oForm.Visible = true;
        }

        public static void ProcesarAsistente(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            int exitosos = 0;
            int fallidos = 0;
            StringBuilder errores = new StringBuilder();

            try
            {
                oForm.Items.Item("OK").Enabled = false;
                oForm.Items.Item("3").Enabled = false;

                string tipo = ((ComboBox)oForm.Items.Item("TIPO").Specific).Selected.Value;
                string operacion = ((ComboBox)oForm.Items.Item("OPERACION").Specific).Selected.Value;
                string periodo = ((ComboBox)oForm.Items.Item("PERIODO").Specific).Selected.Value;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;

                if (string.IsNullOrEmpty(tipo) || string.IsNullOrEmpty(operacion) || string.IsNullOrEmpty(periodo))
                {
                    throw new Exception("Debe seleccionar Tipo, Operación y Período para procesar.");
                }

                int colIndexSel = Globals.GetColIndex(oMatrix, "Sel", 0);
                oForm.Freeze(true);

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    bool seleccionado = ((CheckBox)oMatrix.Columns.Item("Sel").Cells.Item(i).Specific).Checked;
                    if (!seleccionado) continue;

                    string bplId = ((EditText)oMatrix.Columns.Item("BPLId").Cells.Item(i).Specific).Value;
                    string docEntry = ((EditText)oMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific).Value;
                    string bplName = ((EditText)oMatrix.Columns.Item("BPLName").Cells.Item(i).Specific).Value;

                    if (operacion == "-")
                    {
                        if (!string.IsNullOrEmpty(docEntry)) continue;
                    }
                    else if (operacion == "0")
                    {
                        if (string.IsNullOrEmpty(docEntry)) continue;
                    }

                    try
                    {
                        if (tipo == "C") // Compras
                        {
                            if (operacion == "-") // Consulta
                            {
                                string newDocEntry = ProcesarConsultaComprasMasiva(bplId, periodo);
                                ((EditText)oMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific).Value = newDocEntry;
                                oMatrix.CommonSetting.SetCellEditable(i, colIndexSel, false);
                            }
                            else if (operacion == "0") // Reemplazo
                            {
                                if (string.IsNullOrEmpty(docEntry))
                                    throw new Exception("No existe DocEntry para el reemplazo de propuesta de compras.");
                                ProcesarReemplazoComprasMasiva(bplId, periodo, docEntry);
                            }
                        }
                        else if (tipo == "V") // Ventas
                        {
                            if (operacion == "-") // Consulta
                            {
                                string newDocEntry = ProcesarConsultaVentasMasiva(bplId, periodo);
                                ((EditText)oMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific).Value = newDocEntry;
                                oMatrix.CommonSetting.SetCellEditable(i, colIndexSel, false);
                            }
                            else if (operacion == "0") // Reemplazo
                            {
                                if (string.IsNullOrEmpty(docEntry))
                                    throw new Exception("No existe DocEntry para el reemplazo de propuesta de ventas.");
                                ProcesarReemplazoVentasMasiva(bplId, periodo, docEntry);
                            }
                        }
                        exitosos++;
                    }
                    catch (Exception ex)
                    {
                        fallidos++;
                        errores.AppendLine($"Error en fila {i} (Sucursal: {bplId} - {bplName}): {ex.Message}");
                    }
                }

                if (errores.Length > 0)
                {
                    Globals.SBO_Application.MessageBox($"Proceso masivo completado.\nÉxitos: {exitosos}\nFallidos: {fallidos}\n\nDetalle de errores:\n{errores.ToString()}");
                }
                else
                {
                    Globals.SBO_Application.MessageBox($"Proceso masivo completado con éxito.\nTotal procesados: {exitosos}");
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                oForm.Items.Item("3").Enabled = true;
                oForm.Items.Item("OK").Enabled = false;
                Globals.Release(Globals.oRec);
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        #region Compras Masivo

        private static string ProcesarConsultaComprasMasiva(string bplId, string periodo)
        {
            var empresa = Globals.CONF.APIS.FirstOrDefault(x => x.Code == bplId);
            if (empresa == null || string.IsNullOrEmpty(empresa.UEXXAPIS) || string.IsNullOrEmpty(empresa.UEXXUSER) || string.IsNullOrEmpty(empresa.UEXXPASS) || string.IsNullOrEmpty(empresa.UEXXCLID) || string.IsNullOrEmpty(empresa.UEXXCLSE))
                throw new Exception("La empresa seleccionada no tiene la configuración completa de APIs.");

            var documentosSIRE = ConsultaRegistroSIREComprasMasiva(bplId, periodo);
            var documentosSAP = ConsultaRegistroSAPComprasMasiva(bplId, periodo);

            var enAmbos = documentosSAP.Where(sap => documentosSIRE.Any(sire => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sap =>
            {
                sap.UEXXORIGEN = 1;
                return sap;
            }).ToList();

            var soloSAP = documentosSAP.Where(sap => !documentosSIRE.Any(sire => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sap =>
            {
                sap.UEXXORIGEN = 2;
                return sap;
            }).ToList();

            var soloSIRE = documentosSIRE.Where(sire => !documentosSAP.Any(sap => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sire =>
            {
                sire.UEXXORIGEN = 3;
                return sire;
            }).ToList();

            var documentos = enAmbos.Concat(soloSAP).Concat(soloSIRE).ToList();

            SAPbobsCOM.CompanyService oCS = null;
            SAPbobsCOM.GeneralService oGS = null;
            SAPbobsCOM.GeneralData oGD = null;
            SAPbobsCOM.GeneralDataParams oGDP = null;
            try
            {
                oCS = Globals.oCompany.GetCompanyService();
                oGS = oCS.GetGeneralService("EXX_SIRE_COMP");
                oGD = (SAPbobsCOM.GeneralData)oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                oGD.SetProperty("U_EXX_PERIODO", periodo);
                oGD.SetProperty("U_EXX_FECREG", DateTime.Today);
                oGD.SetProperty("U_EXX_BPLID", bplId);
                oGD.SetProperty("U_EXX_ESTADO", "0");

                SAPbobsCOM.GeneralDataCollection oChildren = oGD.Child("EXX_SIRE_COMP1");
                foreach (var doc in documentos.OrderBy(x => x.UEXXORIGEN).ThenBy(y => y.UEXXDOCENTRY).ThenBy(z => z.UEXXFEMI))
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Add();
                    oChild.SetProperty("U_EXX_ORIGEN", doc.UEXXORIGEN);
                    oChild.SetProperty("U_EXX_OBJTYPE", doc.UEXXOBJTYPE ?? "");
                    oChild.SetProperty("U_EXX_DOCENTRY", doc.UEXXDOCENTRY ?? "");
                    oChild.SetProperty("U_EXX_RUC", doc.UEXXRUC ?? "");
                    oChild.SetProperty("U_EXX_RAZSOC", doc.UEXXRAZSOC ?? "");
                    oChild.SetProperty("U_EXX_PERIODO", doc.UEXXPERIODO ?? "");
                    oChild.SetProperty("U_EXX_CARSUN", doc.UEXXCARSUN ?? "");
                    oChild.SetProperty("U_EXX_FEMI", doc.UEXXFEMI ?? "");
                    oChild.SetProperty("U_EXX_FVCTO", doc.UEXXFVCTO ?? "");
                    oChild.SetProperty("U_EXX_TIPDOC", doc.UEXXTIPDOC ?? "");
                    oChild.SetProperty("U_EXX_SERIE", doc.UEXXSERIE ?? "");
                    oChild.SetProperty("U_EXX_ANIO", doc.UEXXANIO ?? "");
                    oChild.SetProperty("U_EXX_NROINI", doc.UEXXNROINI ?? "");
                    oChild.SetProperty("U_EXX_NROFIN", doc.UEXXNROFIN ?? "");
                    oChild.SetProperty("U_EXX_TIPIDE", doc.UEXXTIPIDE ?? "");
                    oChild.SetProperty("U_EXX_NUMIDE", doc.UEXXNUMIDE ?? "");
                    oChild.SetProperty("U_EXX_RAZCLI", doc.UEXXRAZCLI ?? "");

                    oChild.SetProperty("U_EXX_BIGDG", doc.UEXXBIGDG ?? 0.0);
                    oChild.SetProperty("U_EXX_IGVDG", doc.UEXXIGVDG ?? 0.0);
                    oChild.SetProperty("U_EXX_BIGDGN", doc.UEXXBIGDGN ?? 0.0);
                    oChild.SetProperty("U_EXX_IGVDGN", doc.UEXXIGVDGN ?? 0.0);
                    oChild.SetProperty("U_EXX_BIDNG", doc.UEXXBIDNG ?? 0.0);
                    oChild.SetProperty("U_EXX_IGVDNG", doc.UEXXIGVDNG ?? 0.0);
                    oChild.SetProperty("U_EXX_VALNG", doc.UEXXVALNG ?? 0.0);
                    oChild.SetProperty("U_EXX_ISC", doc.UEXXISC ?? 0.0);
                    oChild.SetProperty("U_EXX_ICBPER", doc.UEXXICBPER ?? 0.0);
                    oChild.SetProperty("U_EXX_OTRTRI", doc.UEXXOTRTRI ?? 0.0);
                    oChild.SetProperty("U_EXX_TOTAL", doc.UEXXTOTAL ?? 0.0);

                    oChild.SetProperty("U_EXX_MONEDA", doc.UEXXMONEDA ?? "");
                    oChild.SetProperty("U_EXX_TCAMBIO", doc.UEXXTCAMBIO ?? "");
                    oChild.SetProperty("U_EXX_FEMOD", doc.UEXXFEMOD ?? "");
                    oChild.SetProperty("U_EXX_TIPMOD", doc.UEXXTIPMOD ?? "");
                    oChild.SetProperty("U_EXX_SERMOD", doc.UEXXSERMOD ?? "");
                    oChild.SetProperty("U_EXX_CODDAM", doc.UEXXCODDAM ?? "");
                    oChild.SetProperty("U_EXX_NUMMOD", doc.UEXXNUMMOD ?? "");
                    oChild.SetProperty("U_EXX_CLASIF", doc.UEXXCLASIF ?? "");
                    oChild.SetProperty("U_EXX_IDPROY", doc.UEXXIDPROY ?? "");
                    oChild.SetProperty("U_EXX_PORPAR", doc.UEXXPORPAR ?? 0.0);
                    oChild.SetProperty("U_EXX_IMB", doc.UEXXIMB ?? 0.0);
                    oChild.SetProperty("U_EXX_CARORI", doc.UEXXCARORI ?? "");
                    oChild.SetProperty("U_EXX_DETRA", doc.UEXXDETRA ?? 0.0);
                    oChild.SetProperty("U_EXX_TIPNOT", doc.UEXXTIPNOT ?? "");
                    oChild.SetProperty("U_EXX_ESTCOM", doc.UEXXESTCOM ?? "");
                    oChild.SetProperty("U_EXX_INCAL", doc.UEXXINCAL ?? "");

                    oChild.SetProperty("U_EXX_CLU1", doc.UEXXCLU1 ?? "");
                    oChild.SetProperty("U_EXX_CLU2", doc.UEXXCLU2 ?? "");
                    oChild.SetProperty("U_EXX_CLU3", doc.UEXXCLU3 ?? "");
                    oChild.SetProperty("U_EXX_CLU4", doc.UEXXCLU4 ?? "");
                    oChild.SetProperty("U_EXX_CLU5", doc.UEXXCLU5 ?? "");
                    oChild.SetProperty("U_EXX_CLU6", doc.UEXXCLU6 ?? "");
                    oChild.SetProperty("U_EXX_CLU7", doc.UEXXCLU7 ?? "");
                    oChild.SetProperty("U_EXX_CLU8", doc.UEXXCLU8 ?? "");
                    oChild.SetProperty("U_EXX_CLU9", doc.UEXXCLU9 ?? "");
                    oChild.SetProperty("U_EXX_CLU10", doc.UEXXCLU10 ?? "");
                    oChild.SetProperty("U_EXX_CLU11", doc.UEXXCLU11 ?? "");
                    oChild.SetProperty("U_EXX_CLU12", doc.UEXXCLU12 ?? "");
                    oChild.SetProperty("U_EXX_CLU13", doc.UEXXCLU13 ?? "");
                    oChild.SetProperty("U_EXX_CLU14", doc.UEXXCLU14 ?? "");
                    oChild.SetProperty("U_EXX_CLU15", doc.UEXXCLU15 ?? "");
                    oChild.SetProperty("U_EXX_CLU16", doc.UEXXCLU16 ?? "");
                    oChild.SetProperty("U_EXX_CLU17", doc.UEXXCLU17 ?? "");
                    oChild.SetProperty("U_EXX_CLU18", doc.UEXXCLU18 ?? "");
                    oChild.SetProperty("U_EXX_CLU19", doc.UEXXCLU19 ?? "");
                    oChild.SetProperty("U_EXX_CLU20", doc.UEXXCLU20 ?? "");
                    oChild.SetProperty("U_EXX_CLU21", doc.UEXXCLU21 ?? "");
                    oChild.SetProperty("U_EXX_CLU22", doc.UEXXCLU22 ?? "");
                    oChild.SetProperty("U_EXX_CLU23", doc.UEXXCLU23 ?? "");
                    oChild.SetProperty("U_EXX_CLU24", doc.UEXXCLU24 ?? "");
                    oChild.SetProperty("U_EXX_CLU25", doc.UEXXCLU25 ?? "");
                    oChild.SetProperty("U_EXX_CLU26", doc.UEXXCLU26 ?? "");
                    oChild.SetProperty("U_EXX_CLU27", doc.UEXXCLU27 ?? "");
                    oChild.SetProperty("U_EXX_CLU28", doc.UEXXCLU28 ?? "");
                    oChild.SetProperty("U_EXX_CLU29", doc.UEXXCLU29 ?? "");
                    oChild.SetProperty("U_EXX_CLU30", doc.UEXXCLU30 ?? "");
                    oChild.SetProperty("U_EXX_CLU31", doc.UEXXCLU31 ?? "");
                    oChild.SetProperty("U_EXX_CLU32", doc.UEXXCLU32 ?? "");
                    oChild.SetProperty("U_EXX_CLU33", doc.UEXXCLU33 ?? "");
                    oChild.SetProperty("U_EXX_CLU34", doc.UEXXCLU34 ?? "");
                    oChild.SetProperty("U_EXX_CLU35", doc.UEXXCLU35 ?? "");
                    oChild.SetProperty("U_EXX_CLU36", doc.UEXXCLU36 ?? "");
                    oChild.SetProperty("U_EXX_CLU37", doc.UEXXCLU37 ?? "");
                    oChild.SetProperty("U_EXX_CLU38", doc.UEXXCLU38 ?? "");
                    oChild.SetProperty("U_EXX_CLU39", doc.UEXXCLU39 ?? "");
                }

                oGDP = oGS.Add(oGD);
                return oGDP.GetProperty("DocEntry").ToString();
            }
            finally
            {
                Globals.Release(oCS);
                Globals.Release(oGS);
                Globals.Release(oGD);
                Globals.Release(oGDP);
            }
        }

        private static List<EXXSIRECOMP1> ConsultaRegistroSIREComprasMasiva(string bplId, string periodo)
        {
            try
            {
                List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();
                string periodoSunat = periodo.Replace("-", "");
                string FechaIni = Convert.ToDateTime(periodo + "-01").ToString("yyyy-MM-dd");
                string FechaFin = Convert.ToDateTime(periodo + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");

                string uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/propuesta/{periodoSunat}/exportacioncomprobantepropuesta?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}";
                var propuesta = SIREAPI.ConsultarDocumento<Propuesta>(bplId, uri);

                Thread.Sleep(1000);
                uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodoSunat}&perFin={periodoSunat}&page=1&perPage=20&numTicket={propuesta.numTicket}";
                int intento = 0;
            consultaTicket:
                intento++;
                string fileName = string.Empty;
                byte[] zipBytes = Array.Empty<byte>();
                var ticket = SIREAPI.ConsultarDocumento<Ticket>(bplId, uri);
                if (ticket.Registros[0].ArchivoReporte == null)
                {
                    if (intento > 3)
                    {
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ConsultaExisteLog;
                        else Globals.Query = Properties.Resources.SQL_ConsultaExisteLog;
                        Globals.Query = string.Format(Globals.Query, 0, bplId, FechaIni.Substring(0, 7));
                        Globals.RunQuery(Globals.Query);
                        if (Globals.oRec.RecordCount > 0)
                        {
                            fileName = Globals.oRec.Fields.Item("U_EXX_ARCHTXT").Value.ToString();
                            zipBytes = Convert.FromBase64String(Globals.oRec.Fields.Item("U_EXX_CONTENIDO").Value.ToString());
                            goto TrabajoConLog;
                        }
                        else throw new Exception("No se logró obtener respuesta SUNAT y no existe log guardado para compras.");
                    }
                    goto consultaTicket;
                }
                if (ticket.Registros[0].ArchivoReporte.Count == 0)
                    goto consultaTicket;

                uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte?nomArchivoReporte={ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte}&codTipoAchivoReporte={ticket.Registros[0].ArchivoReporte[0].CodTipoAchivoReporte}&perTributario={periodoSunat}&codProceso={ticket.Registros[0].CodProceso}&numTicket={propuesta.numTicket}&codLibro={Globals.LibroCompra}";
                zipBytes = SIREAPI.ConsultarDocumento<byte[]>(bplId, uri);

            TrabajoConLog:

                #region Consulta excluidos
                //Consulta por API SIRE
                //uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/excluidos/{periodo}/exportaexcluidos?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}&codTipoCDP=01&"; //&codTipoCDP=01";
                uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/excluidos/{periodoSunat}/exportaexcluidos?codTipoArchivo=0&codOrigenEnvio=2"; // &codTipoCDP=01"; //&codTipoCDP=01";
                var excluidos = SIREAPI.ConsultarDocumento<Propuesta>(bplId, uri);

                Thread.Sleep(1000);
                uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodoSunat}&perFin={periodoSunat}&page=1&perPage=20&numTicket={excluidos.numTicket}";
                var ticketExc = SIREAPI.ConsultarDocumento<Ticket>(bplId, uri);

                uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte?nomArchivoReporte={ticketExc.Registros[0].ArchivoReporte[0].NomArchivoReporte}&codTipoAchivoReporte={ticketExc.Registros[0].ArchivoReporte[0].CodTipoAchivoReporte}&perTributario={periodoSunat}&codProceso={ticketExc.Registros[0].CodProceso}&numTicket={excluidos.numTicket}";
                var zipBytesExc = SIREAPI.ConsultarDocumento<byte[]>(bplId, uri);
                #endregion

                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "PROPUESTA", "COMPRAS");

                if (string.IsNullOrEmpty(fileName)) fileName = ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte;
                string fullPath = Path.Combine(folderPath, fileName);

                var zipBytesFinal = UnirRespuestasSireEnBytes(zipBytes, zipBytesExc, fileName.Replace("zip", "txt"));

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                File.WriteAllBytes(fullPath, zipBytes);
                Globals.InformationMessage($"Archivo ZIP de Sunat guardado en: {fullPath}");

                using (var ms = new MemoryStream(zipBytes))
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.Name.EndsWith(".txt"))
                        {
                            using (var entryStream = entry.Open())
                            using (var reader = new StreamReader(entryStream, Encoding.UTF8))
                            {
                                string contenidoTxt = reader.ReadToEnd();
                                list = contenidoTxt.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Skip(1).Where(l => !string.IsNullOrWhiteSpace(l))
                                .Select(linea =>
                                {
                                    var c = linea.Split('|');
                                    return new EXXSIRECOMP1
                                    {
                                        UEXXORIGEN = 3,
                                        UEXXOBJTYPE = string.Empty,
                                        UEXXDOCENTRY = string.Empty,
                                        UEXXRUC = GetString(c, 0),
                                        UEXXRAZSOC = GetString(c, 1),
                                        UEXXPERIODO = GetString(c, 2),
                                        UEXXCARSUN = GetString(c, 3),
                                        UEXXFEMI = GetDateString(c, 4),
                                        UEXXFVCTO = GetDateString(c, 5),
                                        UEXXTIPDOC = GetString(c, 6),
                                        UEXXSERIE = GetString(c, 7),
                                        UEXXANIO = GetString(c, 8),
                                        UEXXNROINI = GetString(c, 9),
                                        UEXXNROFIN = GetString(c, 10),
                                        UEXXTIPIDE = GetString(c, 11),
                                        UEXXNUMIDE = GetString(c, 12),
                                        UEXXRAZCLI = GetString(c, 13),

                                        UEXXBIGDG = GetDouble(c, 14),
                                        UEXXIGVDG = GetDouble(c, 15),
                                        UEXXBIGDGN = GetDouble(c, 16),
                                        UEXXIGVDGN = GetDouble(c, 17),
                                        UEXXBIDNG = GetDouble(c, 18),
                                        UEXXIGVDNG = GetDouble(c, 19),
                                        UEXXVALNG = GetDouble(c, 20),
                                        UEXXISC = GetDouble(c, 21),
                                        UEXXICBPER = GetDouble(c, 22),
                                        UEXXOTRTRI = GetDouble(c, 23),
                                        UEXXTOTAL = GetDouble(c, 24),

                                        UEXXMONEDA = GetString(c, 25),
                                        UEXXTCAMBIO = GetString(c, 26),
                                        UEXXFEMOD = GetDateString(c, 27),
                                        UEXXTIPMOD = GetString(c, 28),
                                        UEXXSERMOD = GetString(c, 29),
                                        UEXXCODDAM = GetString(c, 30),
                                        UEXXNUMMOD = GetString(c, 31),
                                        UEXXCLASIF = GetString(c, 32),
                                        UEXXIDPROY = GetString(c, 33),
                                        UEXXPORPAR = GetDouble(c, 34),
                                        UEXXIMB = GetDouble(c, 35),
                                        UEXXCARORI = GetString(c, 36),
                                        UEXXDETRA = GetDouble(c, 37),
                                        UEXXTIPNOT = GetString(c, 38),
                                        UEXXESTCOM = GetString(c, 39),
                                        UEXXINCAL = GetString(c, 40),
                                        UEXXCLU1 = GetString(c, 41),
                                        UEXXCLU2 = GetString(c, 42),
                                        UEXXCLU3 = GetString(c, 43),
                                        UEXXCLU4 = GetString(c, 44),
                                        UEXXCLU5 = GetString(c, 45),
                                        UEXXCLU6 = GetString(c, 46),
                                        UEXXCLU7 = GetString(c, 47),
                                        UEXXCLU8 = GetString(c, 48),
                                        UEXXCLU9 = GetString(c, 49),
                                        UEXXCLU10 = GetString(c, 50),
                                        UEXXCLU11 = GetString(c, 51),
                                        UEXXCLU12 = GetString(c, 52),
                                        UEXXCLU13 = GetString(c, 53),
                                        UEXXCLU14 = GetString(c, 54),
                                        UEXXCLU15 = GetString(c, 55),
                                        UEXXCLU16 = GetString(c, 56),
                                        UEXXCLU17 = GetString(c, 57),
                                        UEXXCLU18 = GetString(c, 58),
                                        UEXXCLU19 = GetString(c, 59),
                                        UEXXCLU20 = GetString(c, 60),
                                        UEXXCLU21 = GetString(c, 61),
                                        UEXXCLU22 = GetString(c, 62),
                                        UEXXCLU23 = GetString(c, 63),
                                        UEXXCLU24 = GetString(c, 64),
                                        UEXXCLU25 = GetString(c, 65),
                                        UEXXCLU26 = GetString(c, 66),
                                        UEXXCLU27 = GetString(c, 67),
                                        UEXXCLU28 = GetString(c, 68),
                                        UEXXCLU29 = GetString(c, 69),
                                        UEXXCLU30 = GetString(c, 70),
                                        UEXXCLU31 = GetString(c, 71),
                                        UEXXCLU32 = GetString(c, 72),
                                        UEXXCLU33 = GetString(c, 73),
                                        UEXXCLU34 = GetString(c, 74),
                                        UEXXCLU35 = GetString(c, 75),
                                        UEXXCLU36 = GetString(c, 76),
                                        UEXXCLU37 = GetString(c, 77),
                                        UEXXCLU38 = GetString(c, 78),
                                        UEXXCLU39 = GetString(c, 79)
                                    };
                                }).ToList();
                            }
                        }
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("503") && ex.Message.Contains("Unavailable"))
                    throw new Exception("503: El Servicio  no esta disponible por el momento");
                throw ex;
            }
        }

        public static byte[] UnirRespuestasSireEnBytes(byte[] zipBytesPropuesta, byte[] zipBytesExcluidos, string nombreArchivoTxt)
        {
            // 1. Extraer el contenido de texto de ambos ZIPs en memoria
            string contenidoPropuesta = ExtraerTextoDeZipBytes(zipBytesPropuesta);
            string contenidoExcluidos = ExtraerTextoDeZipBytes(zipBytesExcluidos);

            // 2. Retirar la primera fila del archivo de excluidos
            if (!string.IsNullOrEmpty(contenidoExcluidos))
            {
                int posicionSaltoLinea = contenidoExcluidos.IndexOf('\n');

                if (posicionSaltoLinea >= 0)
                {
                    contenidoExcluidos = contenidoExcluidos.Substring(posicionSaltoLinea + 1);
                }
                else
                {
                    // Si solo existe una fila, queda vacío
                    contenidoExcluidos = string.Empty;
                }
            }

            // 3. Concatenar los textos asegurando un salto de línea entre ambos
            StringBuilder sbFinal = new StringBuilder();

            if (!string.IsNullOrWhiteSpace(contenidoPropuesta))
            {
                sbFinal.Append(contenidoPropuesta.TrimEnd());
            }

            if (!string.IsNullOrWhiteSpace(contenidoExcluidos))
            {
                if (sbFinal.Length > 0)
                {
                    sbFinal.AppendLine();
                }

                sbFinal.Append(contenidoExcluidos.TrimEnd());
            }

            string textoUnificado = sbFinal.ToString();

            using (MemoryStream outputZipStream = new MemoryStream())
            {
                using (ZipArchive zipArchive = new ZipArchive(outputZipStream, ZipArchiveMode.Create, true))
                {
                    ZipArchiveEntry entry = zipArchive.CreateEntry(nombreArchivoTxt);

                    using (Stream entryStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(entryStream, Encoding.GetEncoding("ISO-8859-1")))
                    {
                        writer.Write(textoUnificado);
                    }
                }

                return outputZipStream.ToArray();
            }
        }

        private static string ExtraerTextoDeZipBytes(byte[] zipBytes)
        {
            if (zipBytes == null || zipBytes.Length == 0)
                return string.Empty;

            using (MemoryStream ms = new MemoryStream(zipBytes))
            using (ZipArchive archive = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                // SIRE devuelve un único archivo de texto dentro del zip
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
                        entry.FullName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        using (Stream stream = entry.Open())
                        using (StreamReader reader = new StreamReader(stream, Encoding.GetEncoding("ISO-8859-1")))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }
            return string.Empty;
        }


        private static List<EXXSIRECOMP1> ConsultaRegistroSAPComprasMasiva(string bplId, string periodo)
        {
            List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();
            string FechaIni = Convert.ToDateTime(periodo + "-01").ToString("yyyyMMdd");
            string FechaFin = Convert.ToDateTime(periodo + "-01").AddMonths(1).AddDays(-1).ToString("yyyyMMdd");

            if (Globals.IsHana())
                Globals.Query = $"CALL \"SBO_EXX_LE_0801_REGISTRODECOMPRAS_ANEXOS_11\"('{FechaIni}', '{FechaFin}', 'N', '{bplId}')";
            else
                Globals.Query = $"EXEC \"SBO_EXX_LE_0801_REGISTRODECOMPRAS_ANEXOS_11\" '{FechaIni}', '{FechaFin}', 'N', '{bplId}'";

            Globals.RunQuery(Globals.Query);
            Globals.oRec.MoveFirst();

            if (Globals.oRec.RecordCount > 0)
            {
                while (!Globals.oRec.EoF)
                {
                    list.Add(new EXXSIRECOMP1
                    {
                        UEXXORIGEN = 2,
                        UEXXOBJTYPE = Globals.oRec.Fields.Item("ObjType").Value?.ToString(),
                        UEXXDOCENTRY = Globals.oRec.Fields.Item("DocEntry").Value?.ToString(),
                        UEXXRUC = Globals.oRec.Fields.Item("BPE_RUC").Value?.ToString(),
                        UEXXRAZSOC = Globals.oRec.Fields.Item("BPE_ApellidosyNombresoRazonsocial").Value?.ToString(),
                        UEXXPERIODO = Globals.oRec.Fields.Item("BPE_Periodo").Value?.ToString(),
                        UEXXCARSUN = Globals.oRec.Fields.Item("BPE_CARSUNAT").Value?.ToString(),
                        UEXXFEMI = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_Fechadeemision").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_Fechadeemision").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXFVCTO = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_FechaVctoPago").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_FechaVctoPago").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXTIPDOC = Globals.oRec.Fields.Item("BPE_TipoCPDoc").Value?.ToString(),
                        UEXXSERIE = Globals.oRec.Fields.Item("BPE_SeriedelCDP").Value?.ToString(),
                        UEXXANIO = Globals.oRec.Fields.Item("BPE_Ano").Value?.ToString(),

                        UEXXNROINI = Globals.oRec.Fields.Item("BPE_NroCPoDocNroInicialRango").Value?.ToString(),
                        UEXXNROFIN = Globals.oRec.Fields.Item("BPE_NroFinalRango").Value?.ToString(),
                        UEXXTIPIDE = Globals.oRec.Fields.Item("BPE_TipoDocIdentidad").Value?.ToString(),
                        UEXXNUMIDE = Globals.oRec.Fields.Item("BPE_NroDocIdentidad").Value?.ToString(),
                        UEXXRAZCLI = Globals.oRec.Fields.Item("BPE_ApellidosNombresRazonSocial").Value?.ToString(),

                        UEXXBIGDG = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_BIGravadoDG").Value ?? 0.00),
                        UEXXIGVDG = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IGVIPMDG").Value ?? 0.00),
                        UEXXBIGDGN = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_BIGravadoDGNG").Value ?? 0.00),
                        UEXXIGVDGN = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IGVIPMDGNG").Value ?? 0.00),
                        UEXXBIDNG = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_BIGravadoDNG").Value ?? 0.00),
                        UEXXIGVDNG = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IGVIPMDNG").Value ?? 0.00),
                        UEXXVALNG = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ValorAdqNG").Value ?? 0.00),
                        UEXXISC = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ISC").Value ?? 0.00),
                        UEXXICBPER = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ICBPER").Value ?? 0.00),
                        UEXXOTRTRI = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_OtrosTribCargos").Value ?? 0.00),
                        UEXXTOTAL = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_TotalCP").Value ?? 0.00),

                        UEXXMONEDA = Globals.oRec.Fields.Item("BPE_Moneda").Value?.ToString(),
                        UEXXTCAMBIO = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_TipodeCambio").Value?.ToString()) ? "" : Convert.ToDouble(Globals.oRec.Fields.Item("BPE_TipodeCambio").Value ?? 0.000).ToString("0.000"),
                        UEXXFEMOD = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_FechaEmisionDocModificado").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_FechaEmisionDocModificado").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXTIPMOD = Globals.oRec.Fields.Item("BPE_TipoCPModificado").Value?.ToString(),
                        UEXXSERMOD = Globals.oRec.Fields.Item("BPE_SerieCPModificado").Value?.ToString(),
                        UEXXCODDAM = Globals.oRec.Fields.Item("BPE_CODDAMODSI").Value?.ToString(),
                        UEXXNUMMOD = Globals.oRec.Fields.Item("BPE_NroCPModificado").Value?.ToString(),
                        UEXXCLASIF = Globals.oRec.Fields.Item("BPE_ClasifdeBssySss").Value?.ToString(),
                        UEXXIDPROY = Globals.oRec.Fields.Item("BPE_IDProyectoOperadores").Value?.ToString(),
                        UEXXPORPAR = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_PorcPart").Value ?? 0.0),
                        UEXXIMB = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IMB").Value ?? 0.00),
                        UEXXCARORI = Globals.oRec.Fields.Item("BPE_CAROrigIndEoI").Value?.ToString(),
                        UEXXDETRA = 0.0,
                        UEXXTIPNOT = string.Empty,
                        UEXXESTCOM = string.Empty,
                        UEXXINCAL = string.Empty
                    });
                    Globals.oRec.MoveNext();
                }
            }

            return list;
        }

        private static void ProcesarReemplazoComprasMasiva(string bplId, string periodo, string docEntry)
        {
            SAPbobsCOM.CompanyService oCS = null;
            SAPbobsCOM.GeneralService oGS = null;
            SAPbobsCOM.GeneralData oGD = null;
            SAPbobsCOM.GeneralDataParams oGDP = null;
            try
            {
                oCS = Globals.oCompany.GetCompanyService();
                oGS = oCS.GetGeneralService("EXX_SIRE_COMP");

                oGDP = (SAPbobsCOM.GeneralDataParams)oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGDP.SetProperty("DocEntry", docEntry);
                oGD = oGS.GetByParams(oGDP);

                SAPbobsCOM.GeneralDataCollection oChildren = oGD.Child("EXX_SIRE_COMP1");
                List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();

                for (int j = 0; j < oChildren.Count; j++)
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Item(j);
                    list.Add(new EXXSIRECOMP1
                    {
                        UEXXORIGEN = Convert.ToInt32(oChild.GetProperty("U_EXX_ORIGEN")),
                        UEXXRUC = oChild.GetProperty("U_EXX_RUC")?.ToString(),
                        UEXXRAZSOC = oChild.GetProperty("U_EXX_RAZSOC")?.ToString(),
                        UEXXPERIODO = oChild.GetProperty("U_EXX_PERIODO")?.ToString(),
                        UEXXCARSUN = oChild.GetProperty("U_EXX_CARSUN")?.ToString(),
                        UEXXFEMI = oChild.GetProperty("U_EXX_FEMI")?.ToString(),
                        UEXXFVCTO = oChild.GetProperty("U_EXX_FVCTO")?.ToString(),
                        UEXXTIPDOC = oChild.GetProperty("U_EXX_TIPDOC")?.ToString(),
                        UEXXSERIE = oChild.GetProperty("U_EXX_SERIE")?.ToString(),
                        UEXXANIO = oChild.GetProperty("U_EXX_ANIO")?.ToString(),
                        UEXXNROINI = oChild.GetProperty("U_EXX_NROINI")?.ToString(),
                        UEXXNROFIN = oChild.GetProperty("U_EXX_NROFIN")?.ToString(),
                        UEXXTIPIDE = oChild.GetProperty("U_EXX_TIPIDE")?.ToString(),
                        UEXXNUMIDE = oChild.GetProperty("U_EXX_NUMIDE")?.ToString(),
                        UEXXRAZCLI = oChild.GetProperty("U_EXX_RAZCLI")?.ToString(),
                        UEXXBIGDG = Convert.ToDouble(oChild.GetProperty("U_EXX_BIGDG")),
                        UEXXIGVDG = Convert.ToDouble(oChild.GetProperty("U_EXX_IGVDG")),
                        UEXXBIGDGN = Convert.ToDouble(oChild.GetProperty("U_EXX_BIGDGN")),
                        UEXXIGVDGN = Convert.ToDouble(oChild.GetProperty("U_EXX_IGVDGN")),
                        UEXXBIDNG = Convert.ToDouble(oChild.GetProperty("U_EXX_BIDNG")),
                        UEXXIGVDNG = Convert.ToDouble(oChild.GetProperty("U_EXX_IGVDNG")),
                        UEXXVALNG = Convert.ToDouble(oChild.GetProperty("U_EXX_VALNG")),
                        UEXXISC = Convert.ToDouble(oChild.GetProperty("U_EXX_ISC")),
                        UEXXICBPER = Convert.ToDouble(oChild.GetProperty("U_EXX_ICBPER")),
                        UEXXOTRTRI = Convert.ToDouble(oChild.GetProperty("U_EXX_OTRTRI")),
                        UEXXTOTAL = Convert.ToDouble(oChild.GetProperty("U_EXX_TOTAL")),
                        UEXXMONEDA = oChild.GetProperty("U_EXX_MONEDA")?.ToString(),
                        UEXXTCAMBIO = string.IsNullOrEmpty(oChild.GetProperty("U_EXX_TCAMBIO")?.ToString()) ? "" : Convert.ToDouble(oChild.GetProperty("U_EXX_TCAMBIO")) == 0 ? "" : Convert.ToDouble(oChild.GetProperty("U_EXX_TCAMBIO")).ToString("0.000"),
                        UEXXFEMOD = oChild.GetProperty("U_EXX_FEMOD")?.ToString(),
                        UEXXTIPMOD = oChild.GetProperty("U_EXX_TIPMOD")?.ToString(),
                        UEXXSERMOD = oChild.GetProperty("U_EXX_SERMOD")?.ToString(),
                        UEXXCODDAM = oChild.GetProperty("U_EXX_CODDAM")?.ToString(),
                        UEXXNUMMOD = oChild.GetProperty("U_EXX_NUMMOD")?.ToString(),
                        UEXXCLASIF = oChild.GetProperty("U_EXX_CLASIF")?.ToString(),
                        UEXXIDPROY = oChild.GetProperty("U_EXX_IDPROY")?.ToString(),
                        UEXXPORPAR = Convert.ToDouble(oChild.GetProperty("U_EXX_PORPAR")),
                        UEXXIMB = Convert.ToDouble(oChild.GetProperty("U_EXX_IMB")),
                        UEXXCARORI = oChild.GetProperty("U_EXX_CARORI")?.ToString(),
                        UEXXDETRA = Convert.ToDouble(oChild.GetProperty("U_EXX_DETRA")),
                        UEXXTIPNOT = oChild.GetProperty("U_EXX_TIPNOT")?.ToString(),
                        UEXXESTCOM = oChild.GetProperty("U_EXX_ESTCOM")?.ToString(),
                        UEXXINCAL = oChild.GetProperty("U_EXX_INCAL")?.ToString()
                    });
                }

                string ticket = EnviarReemplazoComprasMasivo(bplId, periodo, list);
                ticket += "-" + EnviarNoDomiciliadoComprasMasivo(bplId, periodo);

                oGD.SetProperty("U_EXX_ESTADO", "1");
                oGD.SetProperty("U_EXX_TICKET", ticket);
                oGS.Update(oGD);
            }
            finally
            {
                Globals.Release(oCS);
                Globals.Release(oGS);
                Globals.Release(oGD);
                Globals.Release(oGDP);
            }
        }

        private static string EnviarReemplazoComprasMasivo(string bplId, string periodo, List<EXXSIRECOMP1> list)
        {
            string periodoSunat = periodo.Replace("-", "");
            var empresa = Globals.CONF.APIS.Where(x => x.Code == bplId).ToList();

            string resultado = string.Join(Environment.NewLine,
                list.Where(x => x.UEXXORIGEN != 3).Select(x => string.Join("|", new string[]
                {
                    x.UEXXRUC ?? "",
                    x.UEXXRAZSOC ?? "",
                    x.UEXXPERIODO ?? "",
                    "",
                    Globals.FormatearFecha(x.UEXXFEMI),
                    Globals.FormatearFecha(x.UEXXFVCTO),
                    x.UEXXTIPDOC ?? "",
                    x.UEXXSERIE ?? "",
                    x.UEXXANIO ?? "",
                    x.UEXXNROINI ?? "",
                    x.UEXXNROFIN ?? "",
                    x.UEXXTIPIDE ?? "",
                    x.UEXXNUMIDE ?? "",
                    x.UEXXRAZCLI ?? "",
                    (x.UEXXBIGDG ?? 0).ToString("0.00"),
                    (x.UEXXIGVDG ?? 0).ToString("0.00"),
                    (x.UEXXBIGDGN ?? 0).ToString("0.00"),
                    (x.UEXXIGVDGN ?? 0).ToString("0.00"),
                    (x.UEXXBIDNG ?? 0).ToString("0.00"),
                    (x.UEXXIGVDNG ?? 0).ToString("0.00"),
                    (x.UEXXVALNG ?? 0).ToString("0.00"),
                    (x.UEXXISC ?? 0).ToString("0.00"),
                    (x.UEXXICBPER ?? 0).ToString("0.00"),
                    (x.UEXXOTRTRI ?? 0).ToString("0.00"),
                    (x.UEXXTOTAL ?? 0).ToString("0.00"),
                    x.UEXXMONEDA ?? "",
                    x.UEXXTCAMBIO ?? "",
                    Globals.FormatearFecha(x.UEXXFEMOD),
                    x.UEXXTIPMOD ?? "",
                    x.UEXXSERMOD ?? "",
                    x.UEXXCODDAM ?? "",
                    x.UEXXNUMMOD ?? "",
                    x.UEXXCLASIF ?? "",
                    x.UEXXIDPROY ?? "",
                    (x.UEXXPORPAR ?? 0).ToString("0.00"),
                    (x.UEXXIMB ?? 0).ToString("0.00"),
                    "",
                    ""
                })));

            string fileName = $"LE{empresa[0].UEXXRUC}{periodo.Replace("-", "")}00080400021112";
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "REEMPLAZAR", "COMPRAS");
            string txtPath = Path.Combine(folderPath, $"{fileName}.txt");
            string zipPath = Path.Combine(folderPath, $"{fileName}.zip");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            File.WriteAllText(txtPath, resultado, Encoding.UTF8);
            using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    archive.CreateEntryFromFile(txtPath, Path.GetFileName(txtPath));
                }
            }

            if (File.Exists(txtPath)) File.Delete(txtPath);

            string uri = $"/v1/contribuyente/migeigv/libros/rvierce/receptorpropuesta/web/propuesta/upload";
            string propuesta = SIREAPI.EnviarDocumento(bplId, zipPath, periodoSunat, Globals.LibroCompra, Globals.ProcesoCompra, uri);

            if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir el reemplazo de la propuesta de compras en SIRE.");
            return propuesta;
        }

        private static string EnviarNoDomiciliadoComprasMasivo(string bplId, string periodo)
        {
            string periodoSunat = periodo.Replace("-", "");
            string FechaIni = Convert.ToDateTime(periodo + "-01").ToString("yyyy-MM-dd");
            string FechaFin = Convert.ToDateTime(periodo + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            var empresa = Globals.CONF.APIS.Where(x => x.Code == bplId).ToList();

            string resultado = string.Empty;
            if (Globals.IsHana())
                Globals.Query = $"CALL \"SBO_EXX_LE_0805_REGISTRODECOMPRAS_SIRE\"('{FechaIni.Replace("-", "")}', '{FechaFin.Replace("-", "")}', 'N', '{bplId}')";
            else
                Globals.Query = $"EXEC \"SBO_EXX_LE_0805_REGISTRODECOMPRAS_SIRE\" '{FechaIni.Replace("-", "")}', '{FechaFin.Replace("-", "")}', 'N', '{bplId}'";

            Globals.RunQuery(Globals.Query);
            Globals.oRec.MoveFirst();

            if (Globals.oRec.RecordCount > 0)
            {
                while (!Globals.oRec.EoF)
                {
                    resultado += Globals.oRec.Fields.Item("BPE_Periodo").Value?.ToString() + "|";
                    resultado += "|";
                    resultado += Globals.oRec.Fields.Item("BPE_FechaEmision").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("Indicator").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("Serie").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("Correlativo").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("NoGravada").Value?.ToString() + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("Otro").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Globals.oRec.Fields.Item("DocTotal").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("TipoDocFiscal").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("SerieDocFiscal").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("BPE_AnoDAM").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("CorrelativoDocFiscal").Value?.ToString() + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("RetencionIGV").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Globals.oRec.Fields.Item("Moneda").Value?.ToString() + "|";
                    resultado += string.IsNullOrEmpty(Globals.oRec.Fields.Item("TC").Value?.ToString()) ? "" : Convert.ToDouble(Globals.oRec.Fields.Item("TC").Value?.ToString()) == 0 ? "" : Convert.ToDouble(Globals.oRec.Fields.Item("TC").Value?.ToString()).ToString("0.000") + "|";
                    resultado += Globals.oRec.Fields.Item("Pais").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("NombreBeneficiario").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("Direccion").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("RUC").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_BENEFE").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_NOMBENEFE").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("CodPaisBeneficiario").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_TIPVINECO").Value?.ToString() + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_RENBRU").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_DEDENACAP").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_RENNET").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_TASRET").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_IMPRET").Value?.ToString()).ToString("0.00") + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_CONVENIO").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_EXOOPENDOM").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_TIPREN").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_MODSEREXT").Value?.ToString() + "|";
                    resultado += Globals.oRec.Fields.Item("U_EXX_AP76IR").Value?.ToString() + "|";
                    resultado += "|";
                    resultado += "" + Environment.NewLine;
                    Globals.oRec.MoveNext();
                }
            }

            string fileName = $"LE{empresa[0].UEXXRUC}{periodo.Replace("-", "")}00080500021112";
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "REEMPLAZAR", "COMPRASND");
            string txtPath = Path.Combine(folderPath, $"{fileName}.txt");
            string zipPath = Path.Combine(folderPath, $"{fileName}.zip");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            File.WriteAllText(txtPath, resultado, Encoding.UTF8);
            using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    archive.CreateEntryFromFile(txtPath, Path.GetFileName(txtPath));
                }
            }

            if (File.Exists(txtPath)) File.Delete(txtPath);

            string uri = $"/v1/contribuyente/migeigv/libros/rvierce/receptorpreliminar/web/preliminar/upload";
            string propuesta = SIREAPI.EnviarDocumento(bplId, zipPath, periodoSunat, Globals.LibroNoDom, Globals.ProcesoNoDom, uri);

            if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir de no domiciliados en SIRE.");
            return propuesta;
        }

        #endregion

        #region Ventas Masivo

        private static string ProcesarConsultaVentasMasiva(string bplId, string periodo)
        {
            var empresa = Globals.CONF.APIS.FirstOrDefault(x => x.Code == bplId);
            if (empresa == null || string.IsNullOrEmpty(empresa.UEXXAPIS) || string.IsNullOrEmpty(empresa.UEXXUSER) || string.IsNullOrEmpty(empresa.UEXXPASS) || string.IsNullOrEmpty(empresa.UEXXCLID) || string.IsNullOrEmpty(empresa.UEXXCLSE))
                throw new Exception("La empresa seleccionada no tiene la configuración completa de APIs.");

            var documentosSIRE = ConsultaRegistroSIREVentasMasiva(bplId, periodo);
            var documentosSAP = ConsultaRegistroSAPVentasMasiva(bplId, periodo);

            var enAmbos = documentosSAP.Where(sap => documentosSIRE.Any(sire => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sap =>
            {
                sap.UEXXORIGEN = 1;
                return sap;
            }).ToList();

            var soloSAP = documentosSAP.Where(sap => !documentosSIRE.Any(sire => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sap =>
            {
                sap.UEXXORIGEN = 2;
                return sap;
            }).ToList();

            var soloSIRE = documentosSIRE.Where(sire => !documentosSAP.Any(sap => sap.UEXXRUC == sire.UEXXRUC && sap.UEXXPERIODO == sire.UEXXPERIODO && sap.UEXXTIPDOC == sire.UEXXTIPDOC && sap.UEXXSERIE == sire.UEXXSERIE && sap.UEXXNROINI == sire.UEXXNROINI && sap.UEXXNROFIN == sire.UEXXNROFIN)).Select(sire =>
            {
                sire.UEXXORIGEN = 3;
                return sire;
            }).ToList();

            var documentos = enAmbos.Concat(soloSAP).Concat(soloSIRE).ToList();

            SAPbobsCOM.CompanyService oCS = null;
            SAPbobsCOM.GeneralService oGS = null;
            SAPbobsCOM.GeneralData oGD = null;
            SAPbobsCOM.GeneralDataParams oGDP = null;
            try
            {
                oCS = Globals.oCompany.GetCompanyService();
                oGS = oCS.GetGeneralService("EXX_SIRE_VENT");
                oGD = (SAPbobsCOM.GeneralData)oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                oGD.SetProperty("U_EXX_PERIODO", periodo);
                oGD.SetProperty("U_EXX_FECREG", DateTime.Today);
                oGD.SetProperty("U_EXX_BPLID", bplId);
                oGD.SetProperty("U_EXX_ESTADO", "0");

                SAPbobsCOM.GeneralDataCollection oChildren = oGD.Child("EXX_SIRE_VENT1");
                foreach (var doc in documentos.OrderBy(x => x.UEXXORIGEN).ThenBy(y => y.UEXXDOCENTRY).ThenBy(z => z.UEXXFEMI))
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Add();
                    oChild.SetProperty("U_EXX_ORIGEN", doc.UEXXORIGEN);
                    oChild.SetProperty("U_EXX_OBJTYPE", doc.UEXXOBJTYPE ?? "");
                    oChild.SetProperty("U_EXX_DOCENTRY", doc.UEXXDOCENTRY ?? "");
                    oChild.SetProperty("U_EXX_RUC", doc.UEXXRUC ?? "");
                    oChild.SetProperty("U_EXX_RAZSOC", doc.UEXXRAZSOC ?? "");
                    oChild.SetProperty("U_EXX_PERIODO", doc.UEXXPERIODO ?? "");
                    oChild.SetProperty("U_EXX_CARSUN", doc.UEXXCARSUN ?? "");
                    oChild.SetProperty("U_EXX_FEMI", doc.UEXXFEMI ?? "");
                    oChild.SetProperty("U_EXX_FVCTO", doc.UEXXFVCTO ?? "");
                    oChild.SetProperty("U_EXX_TIPDOC", doc.UEXXTIPDOC ?? "");
                    oChild.SetProperty("U_EXX_SERIE", doc.UEXXSERIE ?? "");
                    oChild.SetProperty("U_EXX_NROINI", doc.UEXXNROINI ?? "");
                    oChild.SetProperty("U_EXX_NROFIN", doc.UEXXNROFIN ?? "");
                    oChild.SetProperty("U_EXX_TIPIDE", doc.UEXXTIPIDE ?? "");
                    oChild.SetProperty("U_EXX_NUMIDE", doc.UEXXNUMIDE ?? "");
                    oChild.SetProperty("U_EXX_RAZCLI", doc.UEXXRAZCLI ?? "");

                    oChild.SetProperty("U_EXX_VFAEX", doc.UEXXVFAEX ?? 0.0);
                    oChild.SetProperty("U_EXX_BIGRA", doc.UEXXBIGRA ?? 0.0);
                    oChild.SetProperty("U_EXX_DESBI", doc.UEXXDESBI ?? 0.0);
                    oChild.SetProperty("U_EXX_IGVIPM", doc.UEXXIGVIPM ?? 0.0);
                    oChild.SetProperty("U_EXX_DESII", doc.UEXXDESII ?? 0.0);
                    oChild.SetProperty("U_EXX_MOEX", doc.UEXXMOEX ?? 0.0);
                    oChild.SetProperty("U_EXX_MOIN", doc.UEXXMOIN ?? 0.0);
                    oChild.SetProperty("U_EXX_ISC", doc.UEXXISC ?? 0.0);
                    oChild.SetProperty("U_EXX_BIGIP", doc.UEXXBIGIP ?? 0.0);
                    oChild.SetProperty("U_EXX_IVAP", doc.UEXXIVAP ?? 0.0);
                    oChild.SetProperty("U_EXX_ICBPER", doc.UEXXICBPER ?? 0.0);
                    oChild.SetProperty("U_EXX_OTRIB", doc.UEXXOTRIB ?? 0.0);
                    oChild.SetProperty("U_EXX_TOTAL", doc.UEXXTOTAL ?? 0.0);

                    oChild.SetProperty("U_EXX_MONEDA", doc.UEXXMONEDA ?? "");
                    oChild.SetProperty("U_EXX_TCAMBIO", doc.UEXXTCAMBIO ?? "");
                    oChild.SetProperty("U_EXX_FEMOD", doc.UEXXFEMOD ?? "");
                    oChild.SetProperty("U_EXX_TIPMOD", doc.UEXXTIPMOD ?? "");
                    oChild.SetProperty("U_EXX_SERMOD", doc.UEXXSERMOD ?? "");
                    oChild.SetProperty("U_EXX_NUMMOD", doc.UEXXNUMMOD ?? "");
                    oChild.SetProperty("U_EXX_IDPOA", doc.UEXXIDPOA ?? "");
                    oChild.SetProperty("U_EXX_TIPNOT", doc.UEXXTIPNOT ?? "");
                    oChild.SetProperty("U_EXX_ESTCOM", doc.UEXXESTCOM ?? "");
                    oChild.SetProperty("U_EXX_VFOBE", doc.UEXXVFOBE ?? 0.0);
                    oChild.SetProperty("U_EXX_VOPGRA", doc.UEXXVOPGRA ?? 0.0);
                    oChild.SetProperty("U_EXX_TIPOPE", doc.UEXXTIPOPE ?? "");
                    oChild.SetProperty("U_EXX_DAMCP", doc.UEXXDAMCP ?? "");
                    oChild.SetProperty("U_EXX_CLU1", doc.UEXXCLU1 ?? "");
                }

                oGDP = oGS.Add(oGD);
                return oGDP.GetProperty("DocEntry").ToString();
            }
            finally
            {
                Globals.Release(oCS);
                Globals.Release(oGS);
                Globals.Release(oGD);
                Globals.Release(oGDP);
            }
        }

        private static List<EXXSIREVENT1> ConsultaRegistroSIREVentasMasiva(string bplId, string periodo)
        {
            List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();
            string periodoSunat = periodo.Replace("-", "");
            string FechaIni = Convert.ToDateTime(periodo + "-01").ToString("yyyy-MM-dd");
            string FechaFin = Convert.ToDateTime(periodo + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");

            string uri = $"/v1/contribuyente/migeigv/libros/rvie/propuesta/web/propuesta/{periodoSunat}/exportapropuesta?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}";
            var propuesta = SIREAPI.ConsultarDocumento<Propuesta>(bplId, uri);

            Thread.Sleep(1000);
            uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodoSunat}&perFin={periodoSunat}&page=1&perPage=20&numTicket={propuesta.numTicket}";
            int intento = 0;
        consultaTicket:
            intento++;
            string fileName = string.Empty;
            byte[] zipBytes = Array.Empty<byte>();
            var ticket = SIREAPI.ConsultarDocumento<Ticket>(bplId, uri);
            if (ticket.Registros[0].ArchivoReporte == null)
            {
                if (intento > 3)
                {
                    if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ConsultaExisteLog;
                    else Globals.Query = Properties.Resources.SQL_ConsultaExisteLog;
                    Globals.Query = string.Format(Globals.Query, 1, bplId, FechaIni.Substring(0, 7));
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.RecordCount > 0)
                    {
                        fileName = Globals.oRec.Fields.Item("U_EXX_ARCHTXT").Value.ToString();
                        zipBytes = Convert.FromBase64String(Globals.oRec.Fields.Item("U_EXX_CONTENIDO").Value.ToString());
                        goto TrabajoConLog;
                    }
                    else throw new Exception("No se logró obtener respuesta SUNAT y no existe log guardado para ventas.");
                }
                goto consultaTicket;
            }
            if (ticket.Registros[0].ArchivoReporte.Count == 0)
                goto consultaTicket;

            uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte?nomArchivoReporte={ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte}&codTipoAchivoReporte={ticket.Registros[0].ArchivoReporte[0].CodTipoAchivoReporte}&perTributario={periodoSunat}&codProceso={ticket.Registros[0].CodProceso}&numTicket={propuesta.numTicket}";
            zipBytes = SIREAPI.ConsultarDocumento<byte[]>(bplId, uri);

        TrabajoConLog:
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "PROPUESTA", "VENTAS");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            if (string.IsNullOrEmpty(fileName)) fileName = $"{DateTime.Now.ToString("yyyyMMdd-HHmmss")}-{ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte}";
            string fullPath = Path.Combine(folderPath, fileName);
            File.WriteAllBytes(fullPath, zipBytes);

            using (var ms = new MemoryStream(zipBytes))
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
            {
                foreach (var entry in zip.Entries)
                {
                    if (entry.Name.EndsWith(".txt") || entry.Name.EndsWith(".csv"))
                    {
                        using (var entryStream = entry.Open())
                        using (var reader = new StreamReader(entryStream, Encoding.UTF8))
                        {
                            string contenidoTxt = reader.ReadToEnd();
                            list = contenidoTxt.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Skip(1).Where(l => !string.IsNullOrWhiteSpace(l))
                            .Select(linea =>
                            {
                                var c = linea.Split('|');
                                return new EXXSIREVENT1
                                {
                                    UEXXORIGEN = 3,
                                    UEXXOBJTYPE = string.Empty,
                                    UEXXDOCENTRY = string.Empty,
                                    UEXXRUC = GetString(c, 0),
                                    UEXXRAZSOC = GetString(c, 1),
                                    UEXXPERIODO = GetString(c, 2),
                                    UEXXCARSUN = GetString(c, 3),
                                    UEXXFEMI = GetDateString(c, 4),
                                    UEXXFVCTO = GetDateString(c, 5),
                                    UEXXTIPDOC = GetString(c, 6),
                                    UEXXSERIE = GetString(c, 7),
                                    UEXXNROINI = GetString(c, 8),
                                    UEXXNROFIN = GetString(c, 9),
                                    UEXXTIPIDE = GetString(c, 10),
                                    UEXXNUMIDE = GetString(c, 11),
                                    UEXXRAZCLI = GetString(c, 12),

                                    UEXXVFAEX = GetDouble(c, 13),
                                    UEXXBIGRA = GetDouble(c, 14),
                                    UEXXDESBI = GetDouble(c, 15),
                                    UEXXIGVIPM = GetDouble(c, 16),
                                    UEXXDESII = GetDouble(c, 17),
                                    UEXXMOEX = GetDouble(c, 18),
                                    UEXXMOIN = GetDouble(c, 19),
                                    UEXXISC = GetDouble(c, 20),
                                    UEXXBIGIP = GetDouble(c, 21),
                                    UEXXIVAP = GetDouble(c, 22),
                                    UEXXICBPER = GetDouble(c, 23),
                                    UEXXOTRIB = GetDouble(c, 24),
                                    UEXXTOTAL = GetDouble(c, 25),

                                    UEXXMONEDA = GetString(c, 26),
                                    UEXXTCAMBIO = GetString(c, 27),
                                    UEXXFEMOD = GetDateString(c, 28),
                                    UEXXTIPMOD = GetString(c, 29),
                                    UEXXSERMOD = GetString(c, 30),
                                    UEXXNUMMOD = GetString(c, 31),
                                    UEXXIDPOA = GetString(c, 32),
                                    UEXXTIPNOT = GetString(c, 33),
                                    UEXXESTCOM = GetString(c, 34),
                                    UEXXVFOBE = GetDouble(c, 35),
                                    UEXXVOPGRA = GetDouble(c, 36),
                                    UEXXTIPOPE = GetString(c, 37),
                                    UEXXDAMCP = GetString(c, 38),
                                    UEXXCLU1 = GetString(c, 39)
                                };
                            }).ToList();
                        }
                    }
                }
            }
            return list;
        }

        private static List<EXXSIREVENT1> ConsultaRegistroSAPVentasMasiva(string bplId, string periodo)
        {
            List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();
            string FechaIni = Convert.ToDateTime(periodo + "-01").ToString("yyyyMMdd");
            string FechaFin = Convert.ToDateTime(periodo + "-01").AddMonths(1).AddDays(-1).ToString("yyyyMMdd");

            if (Globals.IsHana())
                Globals.Query = $"CALL \"SBO_EXX_LL_1401_REGISTRODEVENTAS_REVIE_ANEXOS_3\"('{FechaIni}', '{FechaFin}', 'N', '{bplId}')";
            else
                Globals.Query = $"EXEC \"SBO_EXX_LL_1401_REGISTRODEVENTAS_REVIE_ANEXOS_3\" '{FechaIni}', '{FechaFin}', 'N', '{bplId}'";

            Globals.RunQuery(Globals.Query);
            Globals.oRec.MoveFirst();

            if (Globals.oRec.RecordCount > 0)
            {
                while (!Globals.oRec.EoF)
                {
                    list.Add(new EXXSIREVENT1
                    {
                        UEXXORIGEN = 2,
                        UEXXOBJTYPE = Globals.oRec.Fields.Item("ObjType").Value?.ToString(),
                        UEXXDOCENTRY = Globals.oRec.Fields.Item("DocEntry").Value?.ToString(),
                        UEXXRUC = Globals.oRec.Fields.Item("BPE_Ruc").Value?.ToString(),
                        UEXXRAZSOC = Globals.oRec.Fields.Item("BPE_RazonSocial").Value?.ToString(),
                        UEXXPERIODO = Globals.oRec.Fields.Item("BPE_Periodo").Value?.ToString(),
                        UEXXCARSUN = Globals.oRec.Fields.Item("BPE_CARSUNAT").Value?.ToString(),
                        UEXXFEMI = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_Fechadeemision").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_Fechadeemision").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXFVCTO = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_FechaVctoPago").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_FechaVctoPago").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXTIPDOC = Globals.oRec.Fields.Item("BPE_TipoCPDoc").Value?.ToString(),
                        UEXXSERIE = Globals.oRec.Fields.Item("BPE_SeriedelCDP").Value?.ToString(),
                        UEXXNROINI = Globals.oRec.Fields.Item("BPE_NroCPoDocNroInicialRango").Value?.ToString(),
                        UEXXNROFIN = Globals.oRec.Fields.Item("BPE_NroFinalRango").Value?.ToString(),
                        UEXXTIPIDE = Globals.oRec.Fields.Item("BPE_TipoDocIdentidad").Value?.ToString(),
                        UEXXNUMIDE = Globals.oRec.Fields.Item("BPE_NroDocIdentidad").Value?.ToString(),
                        UEXXRAZCLI = Globals.oRec.Fields.Item("BPE_ApellidosNombresRazonSocial").Value?.ToString(),

                        UEXXVFAEX = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ValorFacturadoExportacion").Value ?? 0.00),
                        UEXXBIGRA = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_BIGravada").Value ?? 0.00),
                        UEXXDESBI = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_DsctoBI").Value ?? 0.00),
                        UEXXIGVIPM = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IGVIPM").Value ?? 0.00),
                        UEXXDESII = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_DsctoIGVIPM").Value ?? 0.00),
                        UEXXMOEX = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_MtoExonerado").Value ?? 0.00),
                        UEXXMOIN = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_MtoInafecto").Value ?? 0.00),
                        UEXXISC = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ISC").Value ?? 0.00),
                        UEXXBIGIP = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_BIGravIVAP").Value ?? 0.00),
                        UEXXIVAP = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IVAP").Value ?? 0.00),
                        UEXXICBPER = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_ICBPER").Value ?? 0.00),
                        UEXXOTRIB = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_OtrosTributos").Value ?? 0.00),
                        UEXXTOTAL = Convert.ToDouble(Globals.oRec.Fields.Item("BPE_TotalCP").Value ?? 0.00),

                        UEXXMONEDA = Globals.oRec.Fields.Item("BPE_Moneda").Value?.ToString(),
                        UEXXTCAMBIO = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_TipoCambio").Value.ToString()) ? "" : Convert.ToDouble(Globals.oRec.Fields.Item("BPE_TipoCambio").Value ?? 0.000).ToString("0.000"),
                        UEXXFEMOD = string.IsNullOrEmpty(Globals.oRec.Fields.Item("BPE_FechaEmisionDocModificado").Value?.ToString()) ? "" : Convert.ToDateTime(Globals.oRec.Fields.Item("BPE_FechaEmisionDocModificado").Value?.ToString()).ToString("yyyyMMdd"),
                        UEXXTIPMOD = Globals.oRec.Fields.Item("BPE_TipoCPModificado").Value?.ToString(),
                        UEXXSERMOD = Globals.oRec.Fields.Item("BPE_SerieCPModificado").Value?.ToString(),
                        UEXXNUMMOD = Globals.oRec.Fields.Item("BPE_NroCPModificado").Value?.ToString(),
                        UEXXIDPOA = Globals.oRec.Fields.Item("BPE_IDProyectoOperadoresAtribucion").Value?.ToString(),
                        UEXXTIPNOT = string.Empty,
                        UEXXESTCOM = string.Empty,

                        UEXXVFOBE = Convert.ToDouble(Globals.oRec.Fields.Item("FOB").Value ?? 0.0),
                        UEXXVOPGRA = 0.0,
                        UEXXTIPOPE = string.Empty,
                        UEXXDAMCP = string.Empty,
                        UEXXCLU1 = string.Empty
                    });

                    Globals.oRec.MoveNext();
                }
            }

            return list;
        }

        private static void ProcesarReemplazoVentasMasiva(string bplId, string periodo, string docEntry)
        {
            SAPbobsCOM.CompanyService oCS = null;
            SAPbobsCOM.GeneralService oGS = null;
            SAPbobsCOM.GeneralData oGD = null;
            SAPbobsCOM.GeneralDataParams oGDP = null;
            try
            {
                oCS = Globals.oCompany.GetCompanyService();
                oGS = oCS.GetGeneralService("EXX_SIRE_VENT");

                oGDP = (SAPbobsCOM.GeneralDataParams)oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGDP.SetProperty("DocEntry", docEntry);
                oGD = oGS.GetByParams(oGDP);

                SAPbobsCOM.GeneralDataCollection oChildren = oGD.Child("EXX_SIRE_VENT1");
                List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();

                for (int j = 0; j < oChildren.Count; j++)
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Item(j);
                    list.Add(new EXXSIREVENT1
                    {
                        UEXXORIGEN = Convert.ToInt32(oChild.GetProperty("U_EXX_ORIGEN")),
                        UEXXRUC = oChild.GetProperty("U_EXX_RUC")?.ToString(),
                        UEXXRAZSOC = oChild.GetProperty("U_EXX_RAZSOC")?.ToString(),
                        UEXXPERIODO = oChild.GetProperty("U_EXX_PERIODO")?.ToString(),
                        UEXXCARSUN = oChild.GetProperty("U_EXX_CARSUN")?.ToString(),
                        UEXXFEMI = oChild.GetProperty("U_EXX_FEMI")?.ToString(),
                        UEXXFVCTO = oChild.GetProperty("U_EXX_FVCTO")?.ToString(),
                        UEXXTIPDOC = oChild.GetProperty("U_EXX_TIPDOC")?.ToString(),
                        UEXXSERIE = oChild.GetProperty("U_EXX_SERIE")?.ToString(),
                        UEXXNROINI = oChild.GetProperty("U_EXX_NROINI")?.ToString(),
                        UEXXNROFIN = oChild.GetProperty("U_EXX_NROFIN")?.ToString(),
                        UEXXTIPIDE = oChild.GetProperty("U_EXX_TIPIDE")?.ToString(),
                        UEXXNUMIDE = oChild.GetProperty("U_EXX_NUMIDE")?.ToString(),
                        UEXXRAZCLI = oChild.GetProperty("U_EXX_RAZCLI")?.ToString(),
                        UEXXVFAEX = Convert.ToDouble(oChild.GetProperty("U_EXX_VFAEX")),
                        UEXXBIGRA = Convert.ToDouble(oChild.GetProperty("U_EXX_BIGRA")),
                        UEXXDESBI = Convert.ToDouble(oChild.GetProperty("U_EXX_DESBI")),
                        UEXXIGVIPM = Convert.ToDouble(oChild.GetProperty("U_EXX_IGVIPM")),
                        UEXXDESII = Convert.ToDouble(oChild.GetProperty("U_EXX_DESII")),
                        UEXXMOEX = Convert.ToDouble(oChild.GetProperty("U_EXX_MOEX")),
                        UEXXMOIN = Convert.ToDouble(oChild.GetProperty("U_EXX_MOIN")),
                        UEXXISC = Convert.ToDouble(oChild.GetProperty("U_EXX_ISC")),
                        UEXXBIGIP = Convert.ToDouble(oChild.GetProperty("U_EXX_BIGIP")),
                        UEXXIVAP = Convert.ToDouble(oChild.GetProperty("U_EXX_IVAP")),
                        UEXXICBPER = Convert.ToDouble(oChild.GetProperty("U_EXX_ICBPER")),
                        UEXXOTRIB = Convert.ToDouble(oChild.GetProperty("U_EXX_OTRIB")),
                        UEXXTOTAL = Convert.ToDouble(oChild.GetProperty("U_EXX_TOTAL")),
                        UEXXMONEDA = oChild.GetProperty("U_EXX_MONEDA")?.ToString(),
                        UEXXTCAMBIO = string.IsNullOrEmpty(oChild.GetProperty("U_EXX_TCAMBIO")?.ToString()) ? "" : Convert.ToDouble(oChild.GetProperty("U_EXX_TCAMBIO")) == 0 ? "" : Convert.ToDouble(oChild.GetProperty("U_EXX_TCAMBIO")).ToString("0.000"),
                        UEXXFEMOD = oChild.GetProperty("U_EXX_FEMOD")?.ToString(),
                        UEXXTIPMOD = oChild.GetProperty("U_EXX_TIPMOD")?.ToString(),
                        UEXXSERMOD = oChild.GetProperty("U_EXX_SERMOD")?.ToString(),
                        UEXXNUMMOD = oChild.GetProperty("U_EXX_NUMMOD")?.ToString(),
                        UEXXIDPOA = oChild.GetProperty("U_EXX_IDPOA")?.ToString(),
                        UEXXTIPNOT = oChild.GetProperty("U_EXX_TIPNOT")?.ToString(),
                        UEXXESTCOM = oChild.GetProperty("U_EXX_ESTCOM")?.ToString(),
                        UEXXVFOBE = Convert.ToDouble(oChild.GetProperty("U_EXX_VFOBE")),
                        UEXXVOPGRA = Convert.ToDouble(oChild.GetProperty("U_EXX_VOPGRA")),
                        UEXXTIPOPE = oChild.GetProperty("U_EXX_TIPOPE")?.ToString(),
                        UEXXDAMCP = oChild.GetProperty("U_EXX_DAMCP")?.ToString(),
                        UEXXCLU1 = oChild.GetProperty("U_EXX_CLU1")?.ToString()
                    });
                }

                string ticket = EnviarReemplazoVentasMasivo(bplId, periodo, list);

                oGD.SetProperty("U_EXX_ESTADO", "1");
                oGD.SetProperty("U_EXX_TICKET", ticket);
                oGS.Update(oGD);
            }
            finally
            {
                Globals.Release(oCS);
                Globals.Release(oGS);
                Globals.Release(oGD);
                Globals.Release(oGDP);
            }
        }

        private static string EnviarReemplazoVentasMasivo(string bplId, string periodo, List<EXXSIREVENT1> list)
        {
            string periodoSunat = periodo.Replace("-", "");
            var empresa = Globals.CONF.APIS.Where(x => x.Code == bplId).ToList();

            string resultado = string.Join(
                Environment.NewLine,
                list.Where(x => x.UEXXORIGEN != 3).Select(x => string.Join("|", new string[]
                {
                    x.UEXXRUC ?? "",
                    x.UEXXRAZSOC ?? "",
                    x.UEXXPERIODO ?? "",
                    x.UEXXCARSUN ?? "",
                    Globals.FormatearFecha(x.UEXXFEMI),
                    Globals.FormatearFecha(x.UEXXFVCTO),
                    x.UEXXTIPDOC ?? "",
                    x.UEXXSERIE ?? "",
                    x.UEXXNROINI ?? "",
                    x.UEXXNROFIN ?? "",
                    x.UEXXTIPIDE ?? "",
                    x.UEXXNUMIDE ?? "",
                    x.UEXXRAZCLI ?? "",
                    (x.UEXXVFAEX ?? 0).ToString("0.00"),
                    (x.UEXXBIGRA ?? 0).ToString("0.00"),
                    (x.UEXXDESBI ?? 0).ToString("0.00"),
                    (x.UEXXIGVIPM ?? 0).ToString("0.00"),
                    (x.UEXXDESII ?? 0).ToString("0.00"),
                    (x.UEXXMOEX ?? 0).ToString("0.00"),
                    (x.UEXXMOIN ?? 0).ToString("0.00"),
                    (x.UEXXISC ?? 0).ToString("0.00"),
                    (x.UEXXBIGIP ?? 0).ToString("0.00"),
                    (x.UEXXIVAP ?? 0).ToString("0.00"),
                    (x.UEXXICBPER ?? 0).ToString("0.00"),
                    (x.UEXXOTRIB ?? 0).ToString("0.00"),
                    (x.UEXXTOTAL ?? 0).ToString("0.00"),
                    x.UEXXMONEDA ?? "",
                    x.UEXXTCAMBIO ?? "",
                    Globals.FormatearFecha(x.UEXXFEMOD),
                    x.UEXXTIPMOD ?? "",
                    x.UEXXSERMOD ?? "",
                    x.UEXXNUMMOD ?? "",
                    x.UEXXIDPOA ?? "",
                    x.UEXXTIPNOT ?? "",
                    x.UEXXESTCOM ?? "",
                    (x.UEXXVFOBE ?? 0).ToString("0.00"),
                    (x.UEXXVOPGRA ?? 0).ToString("0.00"),
                    x.UEXXTIPOPE ?? "",
                    x.UEXXDAMCP ?? "",
                    x.UEXXCLU1 ?? ""
                })));

            string fileName = $"LE{empresa[0].UEXXRUC}{periodo.Replace("-", "")}00140400021112";
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "REEMPLAZAR", "VENTAS");
            string txtPath = Path.Combine(folderPath, $"{fileName}.txt");
            string zipPath = Path.Combine(folderPath, $"{fileName}.zip");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            File.WriteAllText(txtPath, resultado, Encoding.UTF8);
            using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
                {
                    archive.CreateEntryFromFile(txtPath, Path.GetFileName(txtPath));
                }
            }

            if (File.Exists(txtPath)) File.Delete(txtPath);

            string uri = $"/v1/contribuyente/migeigv/libros/rvie/receptorpropuesta/web/propuesta/upload";
            string propuesta = SIREAPI.EnviarDocumento(bplId, zipPath, periodoSunat, Globals.LibroVenta, Globals.ProcesoVenta, uri);

            if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir el reemplazo de la propuesta en SIRE.");
            return propuesta;
        }

        #endregion

        #region Helpers

        private static string GetString(string[] c, int index)
        {
            if (index >= c.Length) return string.Empty;
            return string.IsNullOrWhiteSpace(c[index]) ? string.Empty : c[index].Trim();
        }

        private static double? GetDouble(string[] c, int index)
        {
            if (index >= c.Length) return null;
            if (string.IsNullOrWhiteSpace(c[index])) return null;

            double valor;

            var style = NumberStyles.AllowLeadingSign |
                NumberStyles.AllowDecimalPoint |
                NumberStyles.AllowThousands;

            if (double.TryParse(c[index], style, CultureInfo.InvariantCulture, out valor))
                return valor;

            return null;
        }

        private static string GetDateString(string[] c, int index)
        {
            if (index >= c.Length) return string.Empty;
            if (string.IsNullOrWhiteSpace(c[index])) return string.Empty;

            DateTime fecha;

            return DateTime.TryParseExact(
                c[index],
                "dd/MM/yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out fecha)
                ? fecha.ToString("yyyyMMdd")
                : string.Empty;
        }

        #endregion

        public static void ConsultarEmpresas(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                ComboBox tipo = (ComboBox)oForm.Items.Item("TIPO").Specific;
                ComboBox operacion = (ComboBox)oForm.Items.Item("OPERACION").Specific;
                ComboBox periodo = (ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");

                if (tipo.Selected == null || string.IsNullOrEmpty(tipo.Selected.Value))
                    throw new Exception("Debe seleccionar el tipo de registro que desea ejecutar");
                if (operacion.Selected == null || string.IsNullOrEmpty(tipo.Selected.Value))
                    throw new Exception("Debe seleccionar la operación SIRE que desea ejecutar");
                if (periodo.Selected == null || string.IsNullOrEmpty(tipo.Selected.Value))
                    throw new Exception("Debe seleccionar el período que desea ejecutar");
                if (Globals.IsHana())
                    Globals.Query = $"CALL \"SBO_EXX_SIRE_ASISTENTE_MASIVO\"('{tipo.Selected.Value}','{operacion.Selected.Value}','{periodo.Selected.Value}')";
                else
                    Globals.Query = $"EXEC \"SBO_EXX_SIRE_ASISTENTE_MASIVO\"('{tipo.Selected.Value}','{operacion.Selected.Value}','{periodo.Selected.Value}')";

                oDataTable.ExecuteQuery(Globals.Query);
                oMatrix.Columns.Item("Sel").DataBind.Bind("DT_0", "Sel");
                oMatrix.Columns.Item("BPLId").DataBind.Bind("DT_0", "BPLId");
                oMatrix.Columns.Item("GlblLocNum").DataBind.Bind("DT_0", "GlblLocNum");
                oMatrix.Columns.Item("BPLName").DataBind.Bind("DT_0", "BPLName");
                oMatrix.Columns.Item("U_EXX_APIS").DataBind.Bind("DT_0", "U_EXX_APIS");
                oMatrix.Columns.Item("U_EXX_USER").DataBind.Bind("DT_0", "U_EXX_USER");
                oMatrix.Columns.Item("U_EXX_PASS").DataBind.Bind("DT_0", "U_EXX_PASS");
                oMatrix.Columns.Item("U_EXX_CLID").DataBind.Bind("DT_0", "U_EXX_CLID");
                oMatrix.Columns.Item("U_EXX_CLSE").DataBind.Bind("DT_0", "U_EXX_CLSE");
                oMatrix.Columns.Item("DocEntry").DataBind.Bind("DT_0", "DocEntry");
                oMatrix.LoadFromDataSource();

                int colIndex = Globals.GetColIndex(oMatrix, "Sel", 0);
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    bool seleccionado = ((CheckBox)oMatrix.Columns.Item("Sel").Cells.Item(i).Specific).Checked;
                    if (seleccionado)
                        oMatrix.CommonSetting.SetCellEditable(i, colIndex, false);
                    else
                        oMatrix.CommonSetting.SetCellEditable(i, colIndex, true);
                }
                oForm.Items.Item("OK").Enabled = true;
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }

            finally
            {
                Globals.Release(Globals.oRec);
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void InhabilitarBoton(ItemEvent pVal, Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");

                oDataTable.Rows.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Items.Item("OK").Enabled = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                Globals.Release(Globals.oRec);
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void LinkPressedDinamic(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string tipo = ((SAPbouiCOM.ComboBox)oForm.Items.Item("TIPO").Specific).Selected.Value;

                if (pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                    string LinkedObjectType = tipo == "C" ? "EXX_SIRE_COMP" : "EXX_SIRE_VENT";

                    if (string.IsNullOrEmpty(LinkedObjectType)) BubbleEvent = false;
                    else
                    {
                        SAPbouiCOM.Column oEditColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("DocEntry");
                        SAPbouiCOM.LinkedButton oLink = ((SAPbouiCOM.LinkedButton)(oEditColumn.ExtendedObject));
                        oLink.LinkedObjectType = LinkedObjectType;
                    }
                }

                if (pVal.ActionSuccess)
                {
                    Form oForm2 = Globals.SBO_Application.Forms.ActiveForm;
                    if (tipo == "C")
                    {
                        Modules.Compras.Main.LoadForm(ref oForm2);
                        Modules.Compras.Main.FormOK(ref oForm2);
                    }
                    else
                    {
                        Modules.Ventas.Main.LoadForm(ref oForm2);
                        Modules.Ventas.Main.FormOK(ref oForm2);
                    }
                    oForm2.Refresh();
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}
