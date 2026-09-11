using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using Newtonsoft.Json;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.IO.Compression;
using Form = SAPbouiCOM.Form;
using System.Threading;
using System.Security.Policy;

namespace AddOnConectorSIRE.Modules.Compras
{
    public class Main
    {
        public static void LoadForm(ref Form oForm)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP");
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
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP");
                SAPbouiCOM.DBDataSource oDBDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP1");
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
                    oForm.Items.Item("5").Visible = true;
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
                    oForm.Items.Item("5").Visible = false;
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
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP");
                SAPbouiCOM.DBDataSource oDBDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP1");
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
                    SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP");
                    SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP1");

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
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ValidaExisteRegistroCompra;
                        else Globals.Query = Properties.Resources.SQL_ValidaExisteRegistroCompra;

                        Globals.Query = string.Format(Globals.Query, oDataSource.GetValue("U_EXX_PERIODO", 0), oDataSource.GetValue("U_EXX_BPLID", 0));
                        Globals.RunQuery(Globals.Query);
                        if (Globals.oRec.RecordCount > 0)
                            throw new Exception("Ya existe un registro para el período y empresa seleccionados, primero debe cancelar el anterior para registrar uno nuevo");
                    }
                    else
                    {
                        var existe = SAPSL.Find<EXXSIRECOMP>(new Tuple<string, string, string>("U_EXX_PERIODO", Globals.Operador.Igual, oDataSource.GetValue("U_EXX_PERIODO", 0)),
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

        public static void ProcesarRegistroCompra(ItemEvent pVal, Form oForm, out bool BubbleEvent)
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
                        if (string.IsNullOrEmpty(ArchivoSIRE)) throw new Exception("Primero seleccione el archivo SIRE de compras");
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
                    SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP1");
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
                        oDataSource1.SetValue("U_EXX_ANIO", indice, doc.UEXXANIO);
                        oDataSource1.SetValue("U_EXX_NROINI", indice, doc.UEXXNROINI);
                        oDataSource1.SetValue("U_EXX_NROFIN", indice, doc.UEXXNROFIN);
                        oDataSource1.SetValue("U_EXX_TIPIDE", indice, doc.UEXXTIPIDE);
                        oDataSource1.SetValue("U_EXX_NUMIDE", indice, doc.UEXXNUMIDE);
                        oDataSource1.SetValue("U_EXX_RAZCLI", indice, doc.UEXXRAZCLI);

                        oDataSource1.SetValue("U_EXX_BIGDG", indice, doc.UEXXBIGDG.ToString());
                        oDataSource1.SetValue("U_EXX_IGVDG", indice, doc.UEXXIGVDG.ToString());
                        oDataSource1.SetValue("U_EXX_BIGDGN", indice, doc.UEXXBIGDGN.ToString());
                        oDataSource1.SetValue("U_EXX_IGVDGN", indice, doc.UEXXIGVDGN.ToString());
                        oDataSource1.SetValue("U_EXX_BIDNG", indice, doc.UEXXBIDNG.ToString());
                        oDataSource1.SetValue("U_EXX_IGVDNG", indice, doc.UEXXIGVDNG.ToString());
                        oDataSource1.SetValue("U_EXX_VALNG", indice, doc.UEXXVALNG.ToString());
                        oDataSource1.SetValue("U_EXX_ISC", indice, doc.UEXXISC.ToString());
                        oDataSource1.SetValue("U_EXX_ICBPER", indice, doc.UEXXICBPER.ToString());
                        oDataSource1.SetValue("U_EXX_OTRTRI", indice, doc.UEXXOTRTRI.ToString());
                        oDataSource1.SetValue("U_EXX_TOTAL", indice, doc.UEXXTOTAL.ToString());
                        oDataSource1.SetValue("U_EXX_MONEDA", indice, doc.UEXXMONEDA);
                        oDataSource1.SetValue("U_EXX_TCAMBIO", indice, doc.UEXXTCAMBIO.ToString());

                        oDataSource1.SetValue("U_EXX_FEMOD", indice, doc.UEXXFEMOD);
                        oDataSource1.SetValue("U_EXX_TIPMOD", indice, doc.UEXXTIPMOD);
                        oDataSource1.SetValue("U_EXX_SERMOD", indice, doc.UEXXSERMOD);
                        oDataSource1.SetValue("U_EXX_CODDAM", indice, doc.UEXXCODDAM);
                        oDataSource1.SetValue("U_EXX_NUMMOD", indice, doc.UEXXNUMMOD);
                        oDataSource1.SetValue("U_EXX_CLASIF", indice, doc.UEXXCLASIF);
                        oDataSource1.SetValue("U_EXX_IDPROY", indice, doc.UEXXIDPROY);
                        oDataSource1.SetValue("U_EXX_PORPAR", indice, doc.UEXXPORPAR.ToString());
                        oDataSource1.SetValue("U_EXX_IMB", indice, doc.UEXXIMB.ToString());
                        oDataSource1.SetValue("U_EXX_CARORI", indice, doc.UEXXCARORI);
                        oDataSource1.SetValue("U_EXX_DETRA", indice, doc.UEXXDETRA.ToString());
                        oDataSource1.SetValue("U_EXX_TIPNOT", indice, doc.UEXXTIPNOT);
                        oDataSource1.SetValue("U_EXX_ESTCOM", indice, doc.UEXXESTCOM);
                        oDataSource1.SetValue("U_EXX_INCAL", indice, doc.UEXXINCAL);

                        oDataSource1.SetValue("U_EXX_CLU1", indice, doc.UEXXCLU1);
                        oDataSource1.SetValue("U_EXX_CLU2", indice, doc.UEXXCLU2);
                        oDataSource1.SetValue("U_EXX_CLU3", indice, doc.UEXXCLU3);
                        oDataSource1.SetValue("U_EXX_CLU4", indice, doc.UEXXCLU4);
                        oDataSource1.SetValue("U_EXX_CLU5", indice, doc.UEXXCLU5);
                        oDataSource1.SetValue("U_EXX_CLU6", indice, doc.UEXXCLU6);
                        oDataSource1.SetValue("U_EXX_CLU7", indice, doc.UEXXCLU7);
                        oDataSource1.SetValue("U_EXX_CLU8", indice, doc.UEXXCLU8);
                        oDataSource1.SetValue("U_EXX_CLU9", indice, doc.UEXXCLU9);
                        oDataSource1.SetValue("U_EXX_CLU10", indice, doc.UEXXCLU10);
                        oDataSource1.SetValue("U_EXX_CLU11", indice, doc.UEXXCLU11);
                        oDataSource1.SetValue("U_EXX_CLU12", indice, doc.UEXXCLU12);
                        oDataSource1.SetValue("U_EXX_CLU13", indice, doc.UEXXCLU13);
                        oDataSource1.SetValue("U_EXX_CLU14", indice, doc.UEXXCLU14);
                        oDataSource1.SetValue("U_EXX_CLU15", indice, doc.UEXXCLU15);
                        oDataSource1.SetValue("U_EXX_CLU16", indice, doc.UEXXCLU16);
                        oDataSource1.SetValue("U_EXX_CLU17", indice, doc.UEXXCLU17);
                        oDataSource1.SetValue("U_EXX_CLU18", indice, doc.UEXXCLU18);
                        oDataSource1.SetValue("U_EXX_CLU19", indice, doc.UEXXCLU19);
                        oDataSource1.SetValue("U_EXX_CLU20", indice, doc.UEXXCLU20);
                        oDataSource1.SetValue("U_EXX_CLU21", indice, doc.UEXXCLU21);
                        oDataSource1.SetValue("U_EXX_CLU22", indice, doc.UEXXCLU22);
                        oDataSource1.SetValue("U_EXX_CLU23", indice, doc.UEXXCLU23);
                        oDataSource1.SetValue("U_EXX_CLU24", indice, doc.UEXXCLU24);
                        oDataSource1.SetValue("U_EXX_CLU25", indice, doc.UEXXCLU25);
                        oDataSource1.SetValue("U_EXX_CLU26", indice, doc.UEXXCLU26);
                        oDataSource1.SetValue("U_EXX_CLU27", indice, doc.UEXXCLU27);
                        oDataSource1.SetValue("U_EXX_CLU28", indice, doc.UEXXCLU28);
                        oDataSource1.SetValue("U_EXX_CLU29", indice, doc.UEXXCLU29);
                        oDataSource1.SetValue("U_EXX_CLU30", indice, doc.UEXXCLU30);
                        oDataSource1.SetValue("U_EXX_CLU31", indice, doc.UEXXCLU31);
                        oDataSource1.SetValue("U_EXX_CLU32", indice, doc.UEXXCLU32);
                        oDataSource1.SetValue("U_EXX_CLU33", indice, doc.UEXXCLU33);
                        oDataSource1.SetValue("U_EXX_CLU34", indice, doc.UEXXCLU34);
                        oDataSource1.SetValue("U_EXX_CLU35", indice, doc.UEXXCLU35);
                        oDataSource1.SetValue("U_EXX_CLU36", indice, doc.UEXXCLU36);
                        oDataSource1.SetValue("U_EXX_CLU37", indice, doc.UEXXCLU37);
                        oDataSource1.SetValue("U_EXX_CLU38", indice, doc.UEXXCLU38);
                        oDataSource1.SetValue("U_EXX_CLU39", indice, doc.UEXXCLU39);
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

        public static List<EXXSIRECOMP1> ConsultaRegistroSIRE(ItemEvent pVal, Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;

                List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();

                if (Globals.CONF.UEXXVSIR == "1") //Lee desde archivo
                {
                    string Archivo = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHTXT").Specific).Value;
                    Globals.InformationMessage("Leyendo TXT SIRE");
                    list = File.ReadLines(Archivo).Skip(1).Where(l => !string.IsNullOrWhiteSpace(l))
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
                else
                {
                    Globals.InformationMessage("Obteniendo archivo SIRE de Sunat");
                    string periodo = oPeriodo.Selected.Value.Replace("-", "");
                    string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                    string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                    //Consulta por API SIRE
                    string uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/propuesta/{periodo}/exportacioncomprobantepropuesta?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}"; //&codTipoCDP=01";
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
                            Globals.Query = string.Format(Globals.Query, 0, oEmpresa.Selected.Value.ToString(), FechaIni.Substring(0, 7));
                            Globals.RunQuery(Globals.Query);
                            if (Globals.oRec.RecordCount > 0)
                            {
                                seleccion = Globals.SBO_Application.MessageBox($"Se encontró un log para el período {FechaIni.Substring(0, 7)} consultado el día {Globals.oRec.Fields.Item("U_EXX_DATE").Value.ToString().Substring(0, 10)}. ¿Desea continuar el proceso con este archivo?", 1, "Ok", "Cancel");
                                if (seleccion == 1)
                                {
                                    fileName = Globals.oRec.Fields.Item("U_EXX_ARCHTXT").Value.ToString();
                                    zipBytes = Convert.FromBase64String(Globals.oRec.Fields.Item("U_EXX_CONTENIDO").Value.ToString());
                                    goto TrabajoConLog;
                                }
                                else throw new Exception("Operación cancelada");
                            }
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


                    #region Consulta excluidos
                    //Consulta por API SIRE
                    //uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/excluidos/{periodo}/exportaexcluidos?codTipoArchivo=0&codOrigenEnvio=2&fecEmisionIni={FechaIni}&fecEmisionFin={FechaFin}&codTipoCDP=01&"; //&codTipoCDP=01";
                    uri = $"/v1/contribuyente/migeigv/libros/rce/propuesta/web/excluidos/{periodo}/exportaexcluidos?codTipoArchivo=0&codOrigenEnvio=2"; // &codTipoCDP=01"; //&codTipoCDP=01";
                    var excluidos= SIREAPI.ConsultarDocumento<Propuesta>(oEmpresa.Selected.Value.ToString(), uri);

                    Thread.Sleep(1000);
                    uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={excluidos.numTicket}";
                    var ticketExc = SIREAPI.ConsultarDocumento<Ticket>(oEmpresa.Selected.Value.ToString(), uri);

                    uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/archivoreporte?nomArchivoReporte={ticketExc.Registros[0].ArchivoReporte[0].NomArchivoReporte}&codTipoAchivoReporte={ticketExc.Registros[0].ArchivoReporte[0].CodTipoAchivoReporte}&perTributario={periodo}&codProceso={ticketExc.Registros[0].CodProceso}&numTicket={excluidos.numTicket}";
                    var zipBytesExc = SIREAPI.ConsultarDocumento<byte[]>(oEmpresa.Selected.Value.ToString(), uri);
                    #endregion

                    //Guardar archivo
                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string folderPath = Path.Combine(documentsPath, Globals.AddOnName, Globals.oCompany.CompanyDB, "PROPUESTA", "COMPRAS");
                    if (string.IsNullOrEmpty(fileName)) fileName = ticket.Registros[0].ArchivoReporte[0].NomArchivoReporte;
                    string fullPath = Path.Combine(folderPath, fileName); 
                    
                    var zipBytesFinal = UnirRespuestasSireEnBytes(zipBytes, zipBytesExc, fileName.Replace("zip", "txt")); 

                    if (!Directory.Exists(folderPath))
                        Directory.CreateDirectory(folderPath);

                    File.WriteAllBytes(fullPath, zipBytesFinal);
                    Globals.InformationMessage($"Archivo ZIP de Sunat guardado en: {fullPath}");

                    using (var ms = new MemoryStream(zipBytesFinal))
                    using (var zip = new ZipArchive(ms, ZipArchiveMode.Read))
                    {
                        foreach (var entry in zip.Entries)
                        {
                            // Solo archivos .txt o .csv (según tu caso)
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

        public static List<EXXSIRECOMP1> ConsultaRegistroSAP(ItemEvent pVal, Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;

                List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();
                if (Globals.CONF.UEXXVSAP == "1")
                {
                    string Archivo = ((SAPbouiCOM.EditText)oForm.Items.Item("ARCHSAP").Specific).Value;
                    Globals.InformationMessage("Leyendo TXT SAP");
                    list = File.ReadLines(Archivo).Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(linea =>
                    {
                        var c = linea.Split('|');
                        string ObjType = string.Empty, DocEntry = string.Empty;
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ObtieneLlavesDocumentosCompra;
                        else Globals.Query = Properties.Resources.SQL_ObtieneLlavesDocumentosCompra;
                        Globals.Query = string.Format(Globals.Query, oEmpresa.Selected.Value, GetString(c, 12), GetString(c, 6), GetString(c, 7), GetString(c, 9));  //RUC - TIPO - SERIE - CORRELATIVO
                        Globals.RunQuery(Globals.Query);
                        Globals.oRec.MoveFirst();
                        if (Globals.oRec.RecordCount > 0)
                        {
                            ObjType = Globals.oRec.Fields.Item("ObjType").Value.ToString();
                            DocEntry = Globals.oRec.Fields.Item("DocEntry").Value.ToString();
                        }
                        Globals.Release(Globals.oRec);

                        return new EXXSIRECOMP1
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
                else
                {
                    Globals.InformationMessage("Obteniendo lista de documentos SAP");
                    string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyyMMdd");
                    string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyyMMdd");
                    if (Globals.IsHana())
                        Globals.Query = $"CALL \"SBO_EXX_LE_0801_REGISTRODECOMPRAS_ANEXOS_11\"('{FechaIni}', '{FechaFin}', 'N', '{oEmpresa.Selected.Value}')";
                    else
                        Globals.Query = $"EXEC \"SBO_EXX_LE_0801_REGISTRODECOMPRAS_ANEXOS_11\" '{FechaIni}', '{FechaFin}', 'N', '{oEmpresa.Selected.Value}'";

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
                                UEXXDETRA = 0, //Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_DETRA").Value ?? 0.00),
                                UEXXTIPNOT = string.Empty, //Globals.oRec.Fields.Item("U_EXX_TIPNOT").Value?.ToString(),
                                UEXXESTCOM = string.Empty, //Globals.oRec.Fields.Item("U_EXX_ESTCOM").Value?.ToString(),
                                UEXXINCAL = string.Empty, //Globals.oRec.Fields.Item("U_EXX_INCAL").Value?.ToString(),

                                UEXXCLU1 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU1").Value?.ToString(),
                                UEXXCLU2 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU2").Value?.ToString(),
                                UEXXCLU3 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU3").Value?.ToString(),
                                UEXXCLU4 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU4").Value?.ToString(),
                                UEXXCLU5 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU5").Value?.ToString(),
                                UEXXCLU6 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU6").Value?.ToString(),
                                UEXXCLU7 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU7").Value?.ToString(),
                                UEXXCLU8 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU8").Value?.ToString(),
                                UEXXCLU9 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU9").Value?.ToString(),
                                UEXXCLU10 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU10").Value?.ToString(),
                                UEXXCLU11 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU11").Value?.ToString(),
                                UEXXCLU12 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU12").Value?.ToString(),
                                UEXXCLU13 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU13").Value?.ToString(),
                                UEXXCLU14 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU14").Value?.ToString(),
                                UEXXCLU15 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU15").Value?.ToString(),
                                UEXXCLU16 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU16").Value?.ToString(),
                                UEXXCLU17 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU17").Value?.ToString(),
                                UEXXCLU18 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU18").Value?.ToString(),
                                UEXXCLU19 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU19").Value?.ToString(),
                                UEXXCLU20 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU20").Value?.ToString(),
                                UEXXCLU21 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU21").Value?.ToString(),
                                UEXXCLU22 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU22").Value?.ToString(),
                                UEXXCLU23 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU23").Value?.ToString(),
                                UEXXCLU24 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU24").Value?.ToString(),
                                UEXXCLU25 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU25").Value?.ToString(),
                                UEXXCLU26 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU26").Value?.ToString(),
                                UEXXCLU27 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU27").Value?.ToString(),
                                UEXXCLU28 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU28").Value?.ToString(),
                                UEXXCLU29 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU29").Value?.ToString(),
                                UEXXCLU30 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU30").Value?.ToString(),
                                UEXXCLU31 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU31").Value?.ToString(),
                                UEXXCLU32 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU32").Value?.ToString(),
                                UEXXCLU33 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU33").Value?.ToString(),
                                UEXXCLU34 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU34").Value?.ToString(),
                                UEXXCLU35 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU35").Value?.ToString(),
                                UEXXCLU36 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU36").Value?.ToString(),
                                UEXXCLU37 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU37").Value?.ToString(),
                                UEXXCLU38 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU38").Value?.ToString(),
                                UEXXCLU39 = string.Empty, //Globals.oRec.Fields.Item("U_EXX_CLU39").Value?.ToString()
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

                SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP");
                SAPbouiCOM.DBDataSource oDataSource1 = oForm.DataSources.DBDataSources.Item("@EXX_SIRE_COMP1");

                List<EXXSIRECOMP1> list = new List<EXXSIRECOMP1>();

                for (int i = 0; i < oDataSource1.Size; i++)
                {
                    list.Add(new EXXSIRECOMP1
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
                        UEXXANIO = oDataSource1.GetValue("U_EXX_ANIO", i),

                        UEXXNROINI = oDataSource1.GetValue("U_EXX_NROINI", i),
                        UEXXNROFIN = oDataSource1.GetValue("U_EXX_NROFIN", i),
                        UEXXTIPIDE = oDataSource1.GetValue("U_EXX_TIPIDE", i),
                        UEXXNUMIDE = oDataSource1.GetValue("U_EXX_NUMIDE", i),
                        UEXXRAZCLI = oDataSource1.GetValue("U_EXX_RAZCLI", i),

                        UEXXBIGDG = Convert.ToDouble(oDataSource1.GetValue("U_EXX_BIGDG", i)),
                        UEXXIGVDG = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IGVDG", i)),
                        UEXXBIGDGN = Convert.ToDouble(oDataSource1.GetValue("U_EXX_BIGDGN", i)),
                        UEXXIGVDGN = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IGVDGN", i)),
                        UEXXBIDNG = Convert.ToDouble(oDataSource1.GetValue("U_EXX_BIDNG", i)),
                        UEXXIGVDNG = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IGVDNG", i)),
                        UEXXVALNG = Convert.ToDouble(oDataSource1.GetValue("U_EXX_VALNG", i)),
                        UEXXISC = Convert.ToDouble(oDataSource1.GetValue("U_EXX_ISC", i)),
                        UEXXICBPER = Convert.ToDouble(oDataSource1.GetValue("U_EXX_ICBPER", i)),
                        UEXXOTRTRI = Convert.ToDouble(oDataSource1.GetValue("U_EXX_OTRTRI", i)),
                        UEXXTOTAL = Convert.ToDouble(oDataSource1.GetValue("U_EXX_TOTAL", i)),

                        UEXXMONEDA = oDataSource1.GetValue("U_EXX_MONEDA", i),

                        UEXXTCAMBIO = string.IsNullOrEmpty(oDataSource1.GetValue("U_EXX_TCAMBIO", i).ToString()) ? "" : Convert.ToDouble(oDataSource1.GetValue("U_EXX_TCAMBIO", i)) == 0 ? "" : Convert.ToDouble(oDataSource1.GetValue("U_EXX_TCAMBIO", i)).ToString("0.000"),

                        UEXXFEMOD = oDataSource1.GetValue("U_EXX_FEMOD", i),
                        UEXXTIPMOD = oDataSource1.GetValue("U_EXX_TIPMOD", i),
                        UEXXSERMOD = oDataSource1.GetValue("U_EXX_SERMOD", i),
                        UEXXCODDAM = oDataSource1.GetValue("U_EXX_CODDAM", i),
                        UEXXNUMMOD = oDataSource1.GetValue("U_EXX_NUMMOD", i),
                        UEXXCLASIF = oDataSource1.GetValue("U_EXX_CLASIF", i),
                        UEXXIDPROY = oDataSource1.GetValue("U_EXX_IDPROY", i),

                        UEXXPORPAR = Convert.ToDouble(oDataSource1.GetValue("U_EXX_PORPAR", i)),
                        UEXXIMB = Convert.ToDouble(oDataSource1.GetValue("U_EXX_IMB", i)),

                        UEXXCARORI = oDataSource1.GetValue("U_EXX_CARORI", i),

                        UEXXDETRA = Convert.ToDouble(oDataSource1.GetValue("U_EXX_DETRA", i)),

                        UEXXTIPNOT = oDataSource1.GetValue("U_EXX_TIPNOT", i),
                        UEXXESTCOM = oDataSource1.GetValue("U_EXX_ESTCOM", i),
                        UEXXINCAL = oDataSource1.GetValue("U_EXX_INCAL", i),

                        UEXXCLU1 = oDataSource1.GetValue("U_EXX_CLU1", i)
                    });
                }

                string ticket = EnviarReemplazo(oForm, list);
                ticket += "-" + EnviarNoDomiciliado(oForm);
                ActualizarUDO(oDataSource, ticket);

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

        public static string EnviarReemplazo(Form oForm, List<EXXSIRECOMP1> list)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                string periodo = oPeriodo.Selected.Value.Replace("-", "");
                string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                var empresa = Globals.CONF.APIS.Where(x => x.Code == oEmpresa.Value).ToList();

                string resultado = string.Join(Environment.NewLine,
                list.Where(x => x.UEXXORIGEN != 3).Select(x => string.Join("|", new string[]
                {
                x.UEXXRUC ?? "",
                x.UEXXRAZSOC ?? "",
                x.UEXXPERIODO ?? "",
                "", //x.UEXXCARSUN ?? "",
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
                "", //x.UEXXCARORI ?? "",
                "" //(x.UEXXDETRA ?? 0).ToString("0.00")
                })));

                string fileName = $"LE{empresa[0].UEXXRUC}{oPeriodo.Value.Replace("-", "")}00080400021112";
                //string fileName = $"{empresa[0].UEXXRUC}{Globals.LibroCompra}{oPeriodo.Value.Replace("-", "")}1"; 
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
                        archive.CreateEntryFromFile(
                            txtPath,
                            Path.GetFileName(txtPath)
                        );
                    }
                }

                if (File.Exists(txtPath))
                    File.Delete(txtPath);

                string uri = $"/v1/contribuyente/migeigv/libros/rvierce/receptorpropuesta/web/propuesta/upload";
                string propuesta = SIREAPI.EnviarDocumento(oEmpresa.Selected.Value.ToString(), zipPath, oPeriodo.Value.Replace("-", ""), Globals.LibroCompra, Globals.ProcesoCompra, uri);

                if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir el reemplazo de la propuesta de compras en SIRE.");
                return propuesta;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string EnviarNoDomiciliado(Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oPeriodo = (SAPbouiCOM.ComboBox)oForm.Items.Item("PERIODO").Specific;
                SAPbouiCOM.ComboBox oEmpresa = (SAPbouiCOM.ComboBox)oForm.Items.Item("BPLID").Specific;
                string periodo = oPeriodo.Selected.Value.Replace("-", "");
                string FechaIni = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").ToString("yyyy-MM-dd");
                string FechaFin = Convert.ToDateTime(oPeriodo.Selected.Value + "-01").AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                var empresa = Globals.CONF.APIS.Where(x => x.Code == oEmpresa.Value).ToList();

                string resultado = string.Empty;
                if (Globals.IsHana())
                    Globals.Query = $"CALL \"SBO_EXX_LE_0805_REGISTRODECOMPRAS_SIRE\"('{FechaIni.Replace("-", "")}', '{FechaFin.Replace("-", "")}', 'N', '{oEmpresa.Selected.Value}')";
                else
                    Globals.Query = $"EXEC \"SBO_EXX_LE_0805_REGISTRODECOMPRAS_SIRE\" '{FechaIni.Replace("-", "")}', '{FechaFin.Replace("-", "")}', 'N', '{oEmpresa.Selected.Value}'";

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

                string fileName = $"LE{empresa[0].UEXXRUC}{oPeriodo.Value.Replace("-", "")}00080500021112";
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
                        archive.CreateEntryFromFile(
                            txtPath,
                            Path.GetFileName(txtPath)
                        );
                    }
                }

                if (File.Exists(txtPath))
                    File.Delete(txtPath);

                string uri = $"/v1/contribuyente/migeigv/libros/rvierce/receptorpreliminar/web/preliminar/upload";
                string propuesta = SIREAPI.EnviarDocumento(oEmpresa.Selected.Value.ToString(), zipPath, oPeriodo.Value.Replace("-", ""), Globals.LibroNoDom, Globals.ProcesoNoDom, uri);

                if (string.IsNullOrEmpty(propuesta)) throw new Exception("Ocurrió un error al subir de no domiciliados en SIRE.");
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
                oGS = oCS.GetGeneralService("EXX_SIRE_COMP");

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
                var tickets = oTicket.Value.Split('-').ToList();
                string mensaje = $"{Globals.AddOnName}\n";

                foreach (string numero in tickets)
                {
                    string uri = $"/v1/contribuyente/migeigv/libros/rvierce/gestionprocesosmasivos/web/masivo/consultaestadotickets?perIni={periodo}&perFin={periodo}&page=1&perPage=20&numTicket={numero}";
                    var ticket = SIREAPI.ConsultarDocumento<Ticket>(oEmpresa.Selected.Value.ToString(), uri);

                    if (ticket.Registros != null && ticket.Registros.Count > 0 && ticket.Registros[0].DetalleTicket != null)
                    {
                        mensaje += $"Proceso: {ticket.Registros[0].DesProceso}\n" +
                                    $"Estado: {ticket.Registros[0].DesEstadoProceso}\n" +
                                    $"Archivo: {ticket.Registros[0].NomArchivoImportacion}\n" +
                                    $"Fecha y Hora carga: {Convert.ToDateTime(ticket.Registros[0].DetalleTicket.FecCargaImportacion).ToString("dd/MM/yyyy")} {ticket.Registros[0].DetalleTicket.HoraCargaImportacion}\n" +
                                    $"Filas Informadas: {ticket.Registros[0].DetalleTicket.CntCPInformados}\n" +
                                    $"Filas Validadas: {ticket.Registros[0].DetalleTicket.CntFilasValidada}\n" +
                                    $"Filas con Error: {ticket.Registros[0].DetalleTicket.CntCPError}\n" +
                                    $"---------------------------------\n";
                    }
                    else
                        throw new Exception("No se logró obtener información del ticket");
                }
                Globals.MessageBox(mensaje);
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
