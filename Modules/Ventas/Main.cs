using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using Newtonsoft.Json;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Modules.Ventas
{
    public class Main
    {
        public static void LoadForm(ref Form oForm)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;

                oDataSource.SetValue("UserSign", 0, Globals.oCompany.UserName);
                oDataSource.SetValue("U_EXX_FECREG", 0, Globals.oCompany.GetCompanyDate().ToString("yyyyMMdd"));


                if (Globals.IsHana()) Globals.Query = $"SELECT \"Code\", \"Name\" FROM \"OFPR\" WHERE \"Code\" < '{DateTime.Now.ToString("yyyy-MM")}' ORDER BY \"Code\" DESC";
                else Globals.Query = $"SELECT Code, Name FROM [OFPR] WHERE Code < '{DateTime.Now.ToString("yyyy-MM")}' ORDER BY Code DESC";
                Globals.LlenarCombo(oForm, "PERIODO", Globals.Query, false, "");

                if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ListarEmpresas;
                else Globals.Query = Properties.Resources.SQL_ListarEmpresas;
                Globals.LlenarCombo(oForm, "BPLID", Globals.Query, false, "");

                ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Item.BackColor = 65280;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Item.ForeColor = 1;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Item.BackColor = 42495;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Item.ForeColor = 1;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Item.BackColor = 255;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Item.ForeColor = 16777215;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Caption = "0";
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Caption = "0";
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Caption = "0";

                if (Globals.CONF.UEXXVSIR == "1")
                {
                    oForm.Items.Item("sARCHTXT").Visible = true;
                    oForm.Items.Item("ARCHTXT").Visible = true;
                    oForm.Items.Item("3").Visible = true;
                }
                else
                {
                    oForm.Items.Item("sARCHTXT").Visible = false;
                    oForm.Items.Item("ARCHTXT").Visible = false;
                    oForm.Items.Item("3").Visible = false;
                }

                if (Globals.CONF.UEXXVSAP == "1")
                {
                    oForm.Items.Item("sARCHSAP").Visible = true;
                    oForm.Items.Item("ARCHSAP").Visible = true;
                    oForm.Items.Item("4").Visible = true;
                }
                else
                {
                    oForm.Items.Item("sARCHSAP").Visible = false;
                    oForm.Items.Item("ARCHSAP").Visible = false;
                    oForm.Items.Item("4").Visible = false;
                }

                oForm.Items.Item("sBPLID").Visible = true;
                oForm.Items.Item("BPLID").Visible = true;
                oForm.Items.Item("6").Visible = false;
                oForm.Items.Item("7").Visible = false;

                SetColumnSum(oMatrix, "C_0_16");
                SetColumnSum(oMatrix, "C_0_17");
                SetColumnSum(oMatrix, "C_0_18");
                SetColumnSum(oMatrix, "C_0_19");
                SetColumnSum(oMatrix, "C_0_20");
                SetColumnSum(oMatrix, "C_0_21");
                SetColumnSum(oMatrix, "C_0_22");
                SetColumnSum(oMatrix, "C_0_23");
                SetColumnSum(oMatrix, "C_0_24");
                SetColumnSum(oMatrix, "C_0_25");
                SetColumnSum(oMatrix, "C_0_26");
                SetColumnSum(oMatrix, "C_0_27");
                SetColumnSum(oMatrix, "C_0_28");
                SetColumnSum(oMatrix, "C_0_38");
                SetColumnSum(oMatrix, "C_0_39");

                SAPbouiCOM.CommonSetting oCommonSetting;
                oCommonSetting = oMatrix.CommonSetting;
                oCommonSetting.FixedColumnsCount = 7;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Visible = true;
            }
        }

        public static void FormMode(MenuEvent pVal, ref Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT");
                SAPbouiCOM.DBDataSource oDBDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT1");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                SAPbouiCOM.Button obtTXT = (SAPbouiCOM.Button)oForm.Items.Item("3").Specific;
                SAPbouiCOM.Button obtCargar = (SAPbouiCOM.Button)oForm.Items.Item("4").Specific;
                SAPbouiCOM.EditText oetDocEntry = (SAPbouiCOM.EditText)oForm.Items.Item("0_U_E").Specific;
                SAPbouiCOM.EditText oetFecha = (SAPbouiCOM.EditText)oForm.Items.Item("FECREG").Specific;
                SAPbouiCOM.ComboBox ocbPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox ocbBPLID = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                SAPbouiCOM.EditText oetTXT = (SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific;
                SAPbouiCOM.EditText oetSAP = (SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific;
                SAPbouiCOM.EditText oeTicket = (SAPbouiCOM.EditText)oForm.Items.Item("TICKET").Specific;
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Caption = "0";
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Caption = "0";
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Caption = "0";

                oDBDataSource.SetValue("DocEntry", 0, "");
                oDBDataSource.SetValue("Canceled", 0, "N");

                oDBDataSource.SetValue("U_EXX_BPLID", 0, "");
                oDBDataSource.SetValue("U_EXX_PERIODO", 0, "");
                oDBDataSource.SetValue("U_EXX_ARCHTXT", 0, "");
                oDBDataSource.SetValue("U_EXX_ARCHSAP", 0, "");

                if (pVal.MenuUID == "1282")
                {
                    oDBDataSource.SetValue("U_EXX_FECREG", 0, DateTime.Now.ToString("yyyyMMdd"));
                    obtCargar.Item.Enabled = true;
                    obtTXT.Item.Enabled = true;
                    oetDocEntry.Item.Enabled = false;
                    oetFecha.Item.Enabled = false;
                    ocbPeriodo.Item.Enabled = true;
                    ocbBPLID.Item.Enabled = true;
                    oetTXT.Item.Enabled = false;
                    oetSAP.Item.Enabled = false;
                    oeTicket.Item.Enabled = false;
                    oForm.Items.Item("ESTADO").Enabled = false;
                    oForm.Items.Item("5").Visible = false;
                    oForm.Items.Item("6").Visible = false;
                    oForm.Items.Item("7").Visible = false;

                    if (Globals.CONF.UEXXVSIR == "1")
                    {
                        oForm.Items.Item("sARCHTXT").Visible = true;
                        oForm.Items.Item("ARCHTXT").Visible = true;
                        oForm.Items.Item("3").Visible = true;
                    }
                    else
                    {
                        oForm.Items.Item("sARCHTXT").Visible = false;
                        oForm.Items.Item("ARCHTXT").Visible = false;
                        oForm.Items.Item("3").Visible = false;
                    }

                    if (Globals.CONF.UEXXVSAP == "1")
                    {
                        oForm.Items.Item("sARCHSAP").Visible = true;
                        oForm.Items.Item("ARCHSAP").Visible = true;
                        oForm.Items.Item("4").Visible = true;
                    }
                    else
                    {
                        oForm.Items.Item("sARCHSAP").Visible = false;
                        oForm.Items.Item("ARCHSAP").Visible = false;
                        oForm.Items.Item("4").Visible = false;
                    }

                    if (oDBDataSource1.Size > 0)
                    {
                        oMatrix.FlushToDataSource();
                        oDBDataSource1.Clear();
                        oMatrix.LoadFromDataSource();
                    }
                }
                else
                {
                    obtCargar.Item.Enabled = false;
                    obtTXT.Item.Enabled = false;
                    oetDocEntry.Item.Enabled = true;
                    oetFecha.Item.Enabled = true;
                    ocbPeriodo.Item.Enabled = true;
                    ocbBPLID.Item.Enabled = true;
                    oetTXT.Item.Enabled = true;
                    oetSAP.Item.Enabled = true;
                    oeTicket.Item.Enabled = true;

                    oForm.Items.Item("sARCHTXT").Visible = true;
                    oForm.Items.Item("ARCHTXT").Visible = true;
                    oForm.Items.Item("sARCHSAP").Visible = true;
                    oForm.Items.Item("ARCHSAP").Visible = true;
                    oForm.Items.Item("ESTADO").Enabled = true;
                    oForm.Items.Item("5").Visible = true;
                }

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Refresh();
                GC.Collect();
            }
        }

        public static void FormOK(ref Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT");
                SAPbouiCOM.DBDataSource oDBDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT1");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                SAPbouiCOM.Button obtTXT = (SAPbouiCOM.Button)oForm.Items.Item("3").Specific;
                SAPbouiCOM.Button obtCargar = (SAPbouiCOM.Button)oForm.Items.Item("4").Specific;
                SAPbouiCOM.EditText oetDocEntry = (SAPbouiCOM.EditText)oForm.Items.Item("0_U_E").Specific;
                SAPbouiCOM.EditText oetFecha = (SAPbouiCOM.EditText)oForm.Items.Item("FECREG").Specific;
                SAPbouiCOM.ComboBox ocbPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox ocbBPLID = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                SAPbouiCOM.EditText oetTXT = (SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific;
                SAPbouiCOM.EditText oetSAP = (SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific;

                if (!string.IsNullOrEmpty(oDBDataSource.GetValue("U_EXX_ARCHTXT", 0)))
                {
                    oForm.Items.Item("sARCHTXT").Visible = true;
                    oForm.Items.Item("ARCHTXT").Visible = true;
                    oForm.Items.Item("6").Visible = false;
                }
                else
                {
                    string estado = oDBDataSource.GetValue("U_EXX_ESTADO", 0);
                    oForm.Items.Item("sARCHTXT").Visible = false;
                    oForm.Items.Item("ARCHTXT").Visible = false;

                    string Cancelado = oDBDataSource.GetValue("Canceled", 0);
                    if (Cancelado == "Y")
                    {
                        oForm.Items.Item("6").Visible = false;
                        oForm.Items.Item("7").Visible = false;
                    }
                    else
                    {
                        oForm.Items.Item("6").Visible = estado == "0" ? true : false;
                        oForm.Items.Item("7").Visible = estado == "1" ? true : false;
                    }
                }

                if (!string.IsNullOrEmpty(oDBDataSource.GetValue("U_EXX_ARCHSAP", 0)))
                {
                    oForm.Items.Item("sARCHSAP").Visible = true;
                    oForm.Items.Item("ARCHSAP").Visible = true;
                }
                else
                {
                    oForm.Items.Item("sARCHSAP").Visible = false;
                    oForm.Items.Item("ARCHSAP").Visible = false;
                }

                oForm.Items.Item("3").Visible = false;
                oForm.Items.Item("4").Visible = false;
                oForm.Items.Item("5").Visible = false;

                oetDocEntry.Item.Enabled = false;
                oetFecha.Item.Enabled = false;
                ocbPeriodo.Item.Enabled = false;
                ocbBPLID.Item.Enabled = false;
                oetTXT.Item.Enabled = false;
                oetSAP.Item.Enabled = false;
                oMatrix.Item.Enabled = false;

                var verde = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                var rojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);
                var naranja = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                int enAmbos = 0, soloSAP = 0, soloSIRE = 0;
                for (int i = 0; i < oDBDataSource1.Size; i++)
                {
                    int ORIGEN = Convert.ToInt32(oDBDataSource1.GetValue("U_EXX_ORIGEN", i));
                    if (ORIGEN == 1) enAmbos++; else if (ORIGEN == 2) soloSAP++; else soloSIRE++;
                    oMatrix.CommonSetting.SetRowBackColor(i + 1, ORIGEN == 1 ? verde : ORIGEN == 2 ? naranja : rojo);
                }

                ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Caption = enAmbos.ToString();
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Caption = soloSAP.ToString();
                ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Caption = soloSIRE.ToString();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Refresh();
                GC.Collect();
            }
        }

        public static void SetColumnSum(SAPbouiCOM.Matrix oMatrix, string Columna)
        {
            oMatrix.Columns.Item(Columna).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
        }

        public static void ValidarRegistro(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {
                    SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT");
                    SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT1");

                    if (oDataSource1.Size == 0)
                        throw new Exception("No hay detalles en la lista");
                    else if (oDataSource1.Size == 1)
                    {
                        if (string.IsNullOrEmpty(oDataSource1.GetValue("U_EXX_RUC", 0).ToString()))
                            throw new Exception("No hay detalles en la lista");
                    }

                    if (string.IsNullOrEmpty(oDataSource.GetValue("U_EXX_BPLID", 0).ToString()))
                        throw new Exception("Debe seleccionar una empresa");


                    if (Globals.CONF.UEXXCONS == "1")
                    {
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ValidaExisteRegistroVenta;
                        else Globals.Query = Properties.Resources.SQL_ValidaExisteRegistroVenta;

                        Globals.Query = string.Format(Globals.Query, oDataSource.GetValue("U_EXX_PERIODO", 0), oDataSource.GetValue("U_EXX_BPLID", 0));
                        Globals.RunQuery(Globals.Query);
                        if (Globals.oRec.RecordCount > 0)
                            throw new Exception("Ya existe un registro para el período y empresa seleccionados, primero debe cancelar el anterior para registrar uno nuevo");
                    }
                    else
                    {
                        var existe = SAPSL.Find<EXXSIREVENT>(new Tuple<string, string, string>("U_EXX_PERIODO", Globals.Operador.Igual, oDataSource.GetValue("U_EXX_PERIODO", 0)),
                                                             new Tuple<string, string, string>("U_EXX_BPLID", Globals.Operador.Igual, oDataSource.GetValue("U_EXX_BPLID", 0)),
                                                             new Tuple<string, string, string>("Canceled", Globals.Operador.Igual, "N"));
                        if (existe.Count > 0) throw new Exception("Ya existe un registro para el período y empresa seleccionados, primero debe cancelar el anterior para registrar uno nuevo");
                    }
                }
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

        public static void SeleccionarArchivo(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = false;
            try
            {
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {
                    ClsFolderFileDialog openFileDialog = new ClsFolderFileDialog();
                    string Archivo = openFileDialog.FindFile();
                    if (pVal.ItemUID == "3") ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific).Value = Archivo;
                    if (pVal.ItemUID == "4") ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific).Value = Archivo;
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void ProcesarRegistroVenta(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                {
                    oForm.Freeze(true);
                    string ArchivoSIRE = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific).Value;
                    string ArchivoSAP = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific).Value;
                    SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                    if (oPeriodo.Selected == null) throw new Exception("Debe Seleccionar el período contable");
                    SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                    if (oEmpresa.Selected == null) throw new Exception("Debe Seleccionar una empresa");
                    if (Globals.CONF.UEXXVSIR == "1")
                    {
                        if (string.IsNullOrEmpty(ArchivoSIRE)) throw new Exception("Primero seleccione el archivo SIRE de ventas");
                    }
                    else
                    {
                        var empresa = Globals.CONF.APIS.FirstOrDefault(x => x.Code == oEmpresa.Selected.Value);
                        if (string.IsNullOrEmpty(empresa.UEXXAPIS) || string.IsNullOrEmpty(empresa.UEXXUSER) || string.IsNullOrEmpty(empresa.UEXXPASS) || string.IsNullOrEmpty(empresa.UEXXCLID) || string.IsNullOrEmpty(empresa.UEXXCLSE))
                            throw new Exception("La empresa selccionada no tiene la configuración completa por favor revise en: Gestión > EXX - SIRE SUNAT > EXX - Configuración SIRE");
                    }
                    if (Globals.CONF.UEXXVSAP == "1")
                    {
                        if (string.IsNullOrEmpty(ArchivoSAP)) throw new Exception("Primero seleccione el archivo SAP de compras");
                    }

                    DateTime fecha = Convert.ToDateTime(oPeriodo.Selected.Value + "-01");
                    Globals.InformationMessage("Consultando documentos, por favor espere...");
                    var documentosSIRE = ConsultaRegistroSIRE(pVal, oForm);
                    var documentosSAP = ConsultaRegistroSAP(pVal, oForm);

                    string json1 = JsonConvert.SerializeObject(documentosSIRE);
                    string json2 = JsonConvert.SerializeObject(documentosSAP);

                    Globals.InformationMessage("Comparando listas SAP - SIRE");
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

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                    SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT1");
                    ((SAPbouiCOM.StaticText)oForm.Items.Item("stAmbos").Specific).Caption = enAmbos.Count.ToString();
                    ((SAPbouiCOM.StaticText)oForm.Items.Item("stSAP").Specific).Caption = soloSAP.Count.ToString();
                    ((SAPbouiCOM.StaticText)oForm.Items.Item("stSIRE").Specific).Caption = soloSIRE.Count.ToString();

                    oMatrix.Clear();
                    oDataSource1.Clear();

                    int indice = 0;
                    var verde = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    var rojo = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);
                    var naranja = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                    documentos.OrderBy(x => x.UEXXORIGEN).ThenBy(y => y.UEXXDOCENTRY).ThenBy(z => z.UEXXFEMI).ToList().ForEach(doc =>
                    {
                        if (indice > 0) oMatrix.AddRow();
                        oDataSource1.InsertRecord(indice);
                        oMatrix.FlushToDataSource();

                        oDataSource1.SetValue("U_EXX_OBJTYPE", indice, doc.UEXXOBJTYPE);
                        oDataSource1.SetValue("U_EXX_DOCENTRY", indice, doc.UEXXDOCENTRY);
                        oDataSource1.SetValue("U_EXX_RUC", indice, doc.UEXXRUC);
                        oDataSource1.SetValue("U_EXX_RAZSOC", indice, doc.UEXXRAZSOC);
                        oDataSource1.SetValue("U_EXX_PERIODO", indice, doc.UEXXPERIODO);
                        oDataSource1.SetValue("U_EXX_CARSUN", indice, doc.UEXXCARSUN);
                        oDataSource1.SetValue("U_EXX_FEMI", indice, doc.UEXXFEMI);
                        oDataSource1.SetValue("U_EXX_FVCTO", indice, doc.UEXXFVCTO);
                        oDataSource1.SetValue("U_EXX_TIPDOC", indice, doc.UEXXTIPDOC);
                        oDataSource1.SetValue("U_EXX_SERIE", indice, doc.UEXXSERIE);
                        oDataSource1.SetValue("U_EXX_NROINI", indice, doc.UEXXNROINI);
                        oDataSource1.SetValue("U_EXX_NROFIN", indice, doc.UEXXNROFIN);
                        oDataSource1.SetValue("U_EXX_TIPIDE", indice, doc.UEXXTIPIDE);
                        oDataSource1.SetValue("U_EXX_NUMIDE", indice, doc.UEXXNUMIDE);
                        oDataSource1.SetValue("U_EXX_RAZCLI", indice, doc.UEXXRAZCLI);

                        oDataSource1.SetValue("U_EXX_VFAEX", indice, doc.UEXXVFAEX.ToString());
                        oDataSource1.SetValue("U_EXX_BIGRA", indice, doc.UEXXBIGRA.ToString());
                        oDataSource1.SetValue("U_EXX_DESBI", indice, doc.UEXXDESBI.ToString());
                        oDataSource1.SetValue("U_EXX_IGVIPM", indice, doc.UEXXIGVIPM.ToString());
                        oDataSource1.SetValue("U_EXX_DESII", indice, doc.UEXXDESII.ToString());
                        oDataSource1.SetValue("U_EXX_MOEX", indice, doc.UEXXMOEX.ToString());
                        oDataSource1.SetValue("U_EXX_MOIN", indice, doc.UEXXMOIN.ToString());
                        oDataSource1.SetValue("U_EXX_ISC", indice, doc.UEXXISC.ToString());
                        oDataSource1.SetValue("U_EXX_BIGIP", indice, doc.UEXXICBPER.ToString());
                        oDataSource1.SetValue("U_EXX_IVAP", indice, doc.UEXXIVAP.ToString());
                        oDataSource1.SetValue("U_EXX_ICBPER", indice, doc.UEXXICBPER.ToString());
                        oDataSource1.SetValue("U_EXX_OTRIB", indice, doc.UEXXOTRIB.ToString());
                        oDataSource1.SetValue("U_EXX_TOTAL", indice, doc.UEXXTOTAL.ToString());

                        oDataSource1.SetValue("U_EXX_MONEDA", indice, doc.UEXXMONEDA);
                        oDataSource1.SetValue("U_EXX_TCAMBIO", indice, doc.UEXXTCAMBIO.ToString());

                        oDataSource1.SetValue("U_EXX_FEMOD", indice, doc.UEXXFEMOD);
                        oDataSource1.SetValue("U_EXX_TIPMOD", indice, doc.UEXXTIPMOD);
                        oDataSource1.SetValue("U_EXX_SERMOD", indice, doc.UEXXSERMOD);
                        oDataSource1.SetValue("U_EXX_NUMMOD", indice, doc.UEXXNUMMOD);
                        oDataSource1.SetValue("U_EXX_IDPOA", indice, doc.UEXXIDPOA);
                        oDataSource1.SetValue("U_EXX_TIPNOT", indice, doc.UEXXTIPNOT);
                        oDataSource1.SetValue("U_EXX_ESTCOM", indice, doc.UEXXESTCOM);
                        oDataSource1.SetValue("U_EXX_VFOBE", indice, doc.UEXXVFOBE.ToString());
                        oDataSource1.SetValue("U_EXX_VOPGRA", indice, doc.UEXXVOPGRA.ToString());
                        oDataSource1.SetValue("U_EXX_TIPOPE", indice, doc.UEXXTIPOPE);
                        oDataSource1.SetValue("U_EXX_DAMCP", indice, doc.UEXXDAMCP);
                        oDataSource1.SetValue("U_EXX_CLU1", indice, doc.UEXXCLU1);
                        oDataSource1.SetValue("U_EXX_ORIGEN", indice, doc.UEXXORIGEN.ToString());
                        oMatrix.LoadFromDataSource();
                        oMatrix.CommonSetting.SetRowBackColor(indice + 1, doc.UEXXORIGEN == 1 ? verde : doc.UEXXORIGEN == 2 ? naranja : rojo);
                        indice++;
                    });

                    oMatrix.AutoResizeColumns();
                    Globals.SuccessMessage("La carga se completó correctamente!");
                    Globals.MessageBox("La carga se completó correctamente!");
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static List<EXXSIREVENT1> ConsultaRegistroSIRE(ItemEvent pVal, Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;

                List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();

                if (Globals.CONF.UEXXVSIR == "1") //Lee desde archivo
                {
                    string Archivo = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific).Value;
                    Globals.InformationMessage("Leyendo TXT SIRE");
                    list = File.ReadLines(Archivo).Skip(1).Where(l => !string.IsNullOrWhiteSpace(l))
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
                else
                {
                    Globals.InformationMessage("Obteniendo archivo SIRE de Sunat");
                    string periodo = oPeriodo.Selected.Value.Replace("-", "");
                    string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                    string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                    //Consulta por API SIRE
                    string uri = $"/v1/contribuyente/migeigv/libros/rvie/propuesta/web/propuesta/{periodo}/exportapropuesta?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}"; // &codTipoCDP=01";
                    var propuesta = SIREAPI.ConsultarDocumento<Propuesta>(oEmpresa.Selected.Value.ToString(), uri);

                    Thread.Sleep(1000);
                    uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={propuesta.numTicket}";
                    int intento = 0;
                consultaTicket:
                    intento++;
                    int seleccion = 0;
                    string fileName = string.Empty;
                    byte[] zipBytes = Array.Empty<byte>();
                    var ticket = SIREAPI.ConsultarDocumento<Ticket>(oEmpresa.Selected.Value.ToString(), uri);
                    if (ticket.Registros[0].ArchivoReporte == null)
                    {
                        if (intento > 3)
                        {
                            Globals.InformationMessage("No se logró obtener respuesta SUNAT");
                            if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ConsultaExisteLog;
                            else Globals.Query = Properties.Resources.SQL_ConsultaExisteLog;
                            Globals.Query = string.Format(Globals.Query, 1, oEmpresa.Selected.Value.ToString(), FechaIni.Substring(0, 7));
                            Globals.RunQuery(Globals.Query);
                            if (Globals.oRec.RecordCount > 0)
                            {
                                seleccion = Globals.SBO_Application.MessageBox($"Se encontró un log para el período {FechaIni.Substring(0, 7)} consultado el día {Globals.oRec.Fields.Item("U_EXX_DATE").Value.ToString().Substring(0, 10)}. ¿Desea continuar el proceso con este archivo?", 1, "Ok", "Cancel");
                                if (seleccion == 1)
                                {
                                    fileName = DateTime.Now.ToString("yyyyMMdd - HHmmss") + "-" + Globals.oRec.Fields.Item("U_EXX_ARCHTXT").Value.ToString();
                                    zipBytes = Convert.FromBase64String(Globals.oRec.Fields.Item("U_EXX_CONTENIDO").Value.ToString());
                                    goto TrabajoConLog;
                                }
                                else throw new Exception("Operación cancelada");
                            }
                            else throw new Exception("Operación cancelada");
                        }
                        goto consultaTicket;
                    }
                    if (ticket.Registros[0].ArchivoReporte.Count == 0)
                        goto consultaTicket;

                    if (seleccion == 0)
                    {
                        uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte?nomArchivoReporte={ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte}&codTipoAchivoReporte={ticket.Registros[0].ArchivoReporte[0].CodTipoAchivoReporte}&perTributario={periodo}&codProceso={ticket.Registros[0].CodProceso}&numTicket={propuesta.numTicket}";
                        zipBytes = SIREAPI.ConsultarDocumento<byte[]>(oEmpresa.Selected.Value.ToString(), uri);
                    }

                TrabajoConLog:
                    //Guardar archivo
                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "PROPUESTA", "VENTAS");

                    if (!Directory.Exists(folderPath))
                        Directory.CreateDirectory(folderPath);

                    if (string.IsNullOrEmpty(fileName)) fileName = $"{DateTime.Now.ToString("yyyyMMdd-HHmmss")}-{ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte}";
                    string fullPath = Path.Combine(folderPath, fileName);
                    File.WriteAllBytes(fullPath, zipBytes);
                    Globals.InformationMessage($"Archivo ZIP de Sunat guardado en: {fullPath}");

                    using (var ms = new MemoryStream(zipBytes))
                    using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                    {
                        foreach (var entry in zip.Entries)
                        {
                            // Solo archivos .txt o .csv (según tu caso)
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

        public static List<EXXSIREVENT1> ConsultaRegistroSAP(ItemEvent pVal, Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;

                List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();
                if (Globals.CONF.UEXXVSAP == "1")
                {
                    string Archivo = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific).Value;
                    Globals.InformationMessage("Leyendo TXT SAP");
                    list = File.ReadLines(Archivo).Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(linea =>
                    {
                        var c = linea.Split('|');
                        string ObjType = string.Empty, DocEntry = string.Empty;
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ObtieneLlavesDocumentosVenta;
                        else Globals.Query = Properties.Resources.SQL_ObtieneLlavesDocumentosVenta;
                        Globals.Query = string.Format(Globals.Query, oEmpresa.Selected.Value, GetString(c, 11), GetString(c, 6), GetString(c, 7), Convert.ToDouble(GetString(c, 8)).ToString());  //RUC - TIPO - SERIE - CORRELATIVO
                        Globals.RunQuery(Globals.Query);
                        Globals.oRec.MoveFirst();
                        if (Globals.oRec.RecordCount > 0)
                        {
                            ObjType = Globals.oRec.Fields.Item("ObjType").Value.ToString();
                            DocEntry = Globals.oRec.Fields.Item("DocEntry").Value.ToString();
                        }
                        Globals.Release(Globals.oRec);

                        return new EXXSIREVENT1
                        {
                            UEXXORIGEN = 2,
                            UEXXOBJTYPE = ObjType,
                            UEXXDOCENTRY = DocEntry,
                            UEXXRUC = GetString(c, 0),
                            UEXXRAZSOC = GetString(c, 1),
                            UEXXPERIODO = GetString(c, 2),
                            UEXXCARSUN = GetString(c, 3),
                            UEXXFEMI = GetDateString(c, 4),
                            UEXXFVCTO = GetDateString(c, 5),
                            UEXXTIPDOC = GetString(c, 6),
                            UEXXSERIE = GetString(c, 7),

                            UEXXNROINI = Convert.ToDouble(GetString(c, 8)).ToString(), //GetString(c, 8),
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
                            UEXXESTCOM = string.Empty, //GetString(c, 34),
                            UEXXVFOBE = 0, //GetDouble(c, 35),
                            UEXXVOPGRA = 0, //GetDouble(c, 36),
                            UEXXTIPOPE = string.Empty, //GetString(c, 37),
                            UEXXDAMCP = string.Empty, //GetString(c, 38),
                            UEXXCLU1 = string.Empty, //GetString(c, 39)
                        };
                    }).ToList();
                }
                else
                {
                    Globals.InformationMessage("Obteniendo lista de documentos SAP");
                    string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyyMMdd");
                    string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyyMMdd");
                    if (Globals.IsHana())
                        Globals.Query = $"CALL \"SBO_EXX_LL_1401_REGISTRODEVENTAS_REVIE_ANEXOS_3\"('{FechaIni}', '{FechaFin}', 'N', '{oEmpresa.Selected.Value}')";
                    else
                        Globals.Query = $"EXEC \"SBO_EXX_LL_1401_REGISTRODEVENTAS_REVIE_ANEXOS_3\" '{FechaIni}', '{FechaFin}', 'N', '{oEmpresa.Selected.Value}'";

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
                                UEXXTIPNOT = string.Empty, //Globals.oRec.Fields.Item("BPE_IDProyectoOperadores").Value?.ToString(),
                                UEXXESTCOM = string.Empty, //Globals.oRec.Fields.Item("BPE_IDProyectoOperadores").Value?.ToString(),

                                UEXXVFOBE = Convert.ToDouble(Globals.oRec.Fields.Item("FOB").Value ?? 0.0),
                                UEXXVOPGRA = 0, //Convert.ToDouble(Globals.oRec.Fields.Item("BPE_IMB").Value ?? 0.00),
                                UEXXTIPOPE = string.Empty, //Globals.oRec.Fields.Item("BPE_IDProyectoOperadores").Value?.ToString(),
                                UEXXDAMCP = string.Empty, //Globals.oRec.Fields.Item("BPE_IDProyectoOperadores").Value?.ToString(),
                                UEXXCLU1 = string.Empty //Globals.oRec.Fields.Item("U_EXX_CLU1").Value?.ToString(),
                            });

                            Globals.oRec.MoveNext();
                        }
                    }
                }
                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
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

        public static void ReemplazarPropuesta(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (oForm.Mode != BoFormMode.fm_OK_MODE) throw new Exception("Para enviar el reemplazo no debe haber nada por actualizar en el formulario");

                int envioSire = Globals.SBO_Application.MessageBox("Se enviará el reemplazo de la propuesta a Sunat ¿Desea continuar?", 1, "&Si", "&No");
                if (envioSire != 1) return;
                SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT");
                SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_VENT1");

                List<EXXSIREVENT1> list = new List<EXXSIREVENT1>();

                for (int i = 0; i < oDataSource1.Size; i++)
                {
                    list.Add(new EXXSIREVENT1
                    {
                        UEXXORIGEN = Convert.ToInt32(oDataSource1.GetValue("U_EXX_ORIGEN", i)),
                        UEXXRUC = oDataSource1.GetValue("U_EXX_RUC", i),
                        UEXXRAZSOC = oDataSource1.GetValue("U_EXX_RAZSOC", i),
                        UEXXPERIODO = oDataSource1.GetValue("U_EXX_PERIODO", i),
                        UEXXCARSUN = oDataSource1.GetValue("U_EXX_CARSUN", i),
                        UEXXFEMI = oDataSource1.GetValue("U_EXX_FEMI", i),
                        UEXXFVCTO = oDataSource1.GetValue("U_EXX_FVCTO", i),
                        UEXXTIPDOC = oDataSource1.GetValue("U_EXX_TIPDOC", i),
                        UEXXSERIE = oDataSource1.GetValue("U_EXX_SERIE", i),

                        UEXXNROINI = oDataSource1.GetValue("U_EXX_NROINI", i),
                        UEXXNROFIN = oDataSource1.GetValue("U_EXX_NROFIN", i),
                        UEXXTIPIDE = oDataSource1.GetValue("U_EXX_TIPIDE", i),
                        UEXXNUMIDE = oDataSource1.GetValue("U_EXX_NUMIDE", i),
                        UEXXRAZCLI = oDataSource1.GetValue("U_EXX_RAZCLI", i),

                        UEXXVFAEX = Convert.ToDouble(oDataSource1.GetValue("U_EXX_VFAEX", i)),
                        UEXXBIGRA = Convert.ToDouble(oDataSource1.GetValue("U_EXX_BIGRA", i)),
                        UEXXDESBI = Convert.ToDouble(oDataSource1.GetValue("U_EXX_DESBI", i)),
                        UEXXIGVIPM = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IGVIPM", i)),
                        UEXXDESII = Convert.ToDouble(oDataSource1.GetValue("U_EXX_DESII", i)),
                        UEXXMOEX = Convert.ToDouble(oDataSource1.GetValue("U_EXX_MOEX", i)),
                        UEXXMOIN = Convert.ToDouble(oDataSource1.GetValue("U_EXX_MOIN", i)),
                        UEXXISC = Convert.ToDouble(oDataSource1.GetValue("U_EXX_ISC", i)),
                        UEXXBIGIP = Convert.ToDouble(oDataSource1.GetValue("U_EXX_BIGIP", i)),
                        UEXXIVAP = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IVAP", i)),
                        UEXXICBPER = Convert.ToDouble(oDataSource1.GetValue("U_EXX_ICBPER", i)),
                        UEXXOTRIB = Convert.ToDouble(oDataSource1.GetValue("U_EXX_OTRIB", i)),
                        UEXXTOTAL = Convert.ToDouble(oDataSource1.GetValue("U_EXX_TOTAL", i)),

                        UEXXMONEDA = oDataSource1.GetValue("U_EXX_MONEDA", i),
                        UEXXTCAMBIO = string.IsNullOrEmpty(oDataSource1.GetValue("U_EXX_TCAMBIO", i).ToString()) ? "" : Convert.ToDouble(oDataSource1.GetValue("U_EXX_TCAMBIO", i)) == 0 ? "" : Convert.ToDouble(oDataSource1.GetValue("U_EXX_TCAMBIO", i)).ToString("0.000"),
                        UEXXFEMOD = oDataSource1.GetValue("U_EXX_FEMOD", i),
                        UEXXTIPMOD = oDataSource1.GetValue("U_EXX_TIPMOD", i),
                        UEXXSERMOD = oDataSource1.GetValue("U_EXX_SERMOD", i),
                        UEXXNUMMOD = oDataSource1.GetValue("U_EXX_NUMMOD", i),
                        UEXXIDPOA = oDataSource1.GetValue("U_EXX_IDPOA", i),
                        UEXXTIPNOT = oDataSource1.GetValue("U_EXX_TIPNOT", i),
                        UEXXESTCOM = oDataSource1.GetValue("U_EXX_ESTCOM", i),
                        UEXXVFOBE = Convert.ToDouble(oDataSource1.GetValue("U_EXX_VFOBE", i)),
                        UEXXVOPGRA = Convert.ToDouble(oDataSource1.GetValue("U_EXX_VOPGRA", i)),
                        UEXXTIPOPE = oDataSource1.GetValue("U_EXX_TIPOPE", i),
                        UEXXDAMCP = oDataSource1.GetValue("U_EXX_DAMCP", i),
                        UEXXCLU1 = oDataSource1.GetValue("U_EXX_CLU1", i)
                    });
                }

                string Ticket = EnviarReemplazo(oForm, list);
                ActualizarUDO(oDataSource1, Ticket);

                if (Globals.SBO_Application.Menus.Item("1304").Enabled)
                    Globals.SBO_Application.Menus.Item("1304").Activate();

                Globals.SuccessMessage("Envío de reemplazo se completó correctamente!");
                Globals.MessageBox("Envío de reemplazo se completó correctamente!");


            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static string EnviarReemplazo(Form oForm, List<EXXSIREVENT1> list)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                string periodo = oPeriodo.Selected.Value.Replace("-", "");
                string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                var empresa = Globals.CONF.APIS.Where(x => x.Code == oEmpresa.Value).ToList();

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
                    ""
                    //x.UEXXTIPNOT ?? "",
                    //x.UEXXESTCOM ?? "",

                    //(x.UEXXVFOBE ?? 0).ToString("0.00"),
                    //(x.UEXXVOPGRA ?? 0).ToString("0.00"),

                    //x.UEXXTIPOPE ?? "",
                    //x.UEXXDAMCP ?? "",
                    //x.UEXXCLU1 ?? ""
                })));

                string fileName = $"LE{empresa[0].UEXXRUC}{oPeriodo.Value.Replace("-", "")}00140400021112";
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
                        archive.CreateEntryFromFile(
                            txtPath,
                            Path.GetFileName(txtPath)
                        );
                    }
                }

                if (File.Exists(txtPath))
                    File.Delete(txtPath);

                string uri = $"/v1/contribuyente/migeigv/libros/rvierce/receptorpropuesta/web/propuesta/upload";
                string propuesta = SIREAPI.EnviarDocumento(oEmpresa.Selected.Value.ToString(), zipPath, oPeriodo.Value.Replace("-", ""), Globals.LibroVenta, Globals.ProcesoVenta, uri);

                if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir el reemplazo de la propuesta en SIRE.");
                return propuesta;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ActualizarUDO(DBDataSource oDataSource, string Ticket)
        {
            SAPbobsCOM.CompanyService oCS = null;
            SAPbobsCOM.GeneralService oGS = null;
            SAPbobsCOM.GeneralData oGD = null;
            SAPbobsCOM.GeneralDataParams oGDP = null;
            try
            {
                string DocEntry = oDataSource.GetValue("DocEntry", 0);
                oCS = Globals.oCompany.GetCompanyService();
                oGS = oCS.GetGeneralService("EXX_SIRE_VENT");

                oGDP = (SAPbobsCOM.GeneralDataParams)oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGDP.SetProperty("DocEntry", DocEntry);
                oGD = oGS.GetByParams(oGDP);
                oGD.SetProperty("U_EXX_ESTADO", "1");
                oGD.SetProperty("U_EXX_TICKET", Ticket);
                oGS.Update(oGD);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oCS);
                Globals.Release(oGS);
                Globals.Release(oGDP);
            }
        }

        public static void ConsultaTicket(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.EditText oTicket = (SAPbouiCOM.EditText)oForm.Items.Item("TICKET").Specific;
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                string periodo = oPeriodo.Selected.Value.Replace("-", "");
                string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");

                string uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={oTicket.Value}";
                var ticket = SIREAPI.ConsultarDocumento<Ticket>(oEmpresa.Selected.Value.ToString(), uri);

                if (ticket.Registros != null && ticket.Registros.Count > 0 && ticket.Registros[0].DetalleTicket != null)
                {
                    Globals.MessageBox($"{Globals.AddOnName}\n" +
                                       $"Proceso: {ticket.Registros[0].DesProceso}\n" +
                                       $"Estado: {ticket.Registros[0].DesEstadoProceso}\n" +
                                       $"Archivo: {ticket.Registros[0].NomArchivoImportacion}\n" +
                                       $"Fecha y Hora carga: {Convert.ToDateTime(ticket.Registros[0].DetalleTicket.FecCargaImportacion).ToString("dd/MM/yyyy")} {ticket.Registros[0].DetalleTicket.HoraCargaImportacion}\n" +
                                       $"Filas Informadas: {ticket.Registros[0].DetalleTicket.CntCPInformados}\n" +
                                       $"Filas Validadas: {ticket.Registros[0].DetalleTicket.CntFilasValidada}\n" +
                                       $"Filas con Error: {ticket.Registros[0].DetalleTicket.CntCPError}\n" +
                                       $"");
                }
                else
                    throw new Exception("No se logró obtener información del ticket");
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void SeleccionarFila(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                if (pVal.Row > 0 && pVal.Row <= oMatrix.RowCount) oMatrix.SelectRow(pVal.Row, true, false);
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void LinkPressedDinamic(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                    string LinkedObjectType = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific).Value.ToString().Trim();

                    if (string.IsNullOrEmpty(LinkedObjectType)) BubbleEvent = false;
                    else
                    {
                        SAPbouiCOM.Column oEditColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("C_0_2");
                        SAPbouiCOM.LinkedButton oLink = ((SAPbouiCOM.LinkedButton)(oEditColumn.ExtendedObject));
                        oLink.LinkedObjectType = LinkedObjectType;
                    }
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
