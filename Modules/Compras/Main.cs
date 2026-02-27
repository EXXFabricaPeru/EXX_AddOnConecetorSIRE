using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using Newtonsoft.Json;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Form = SAPbouiCOM.Form;

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
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.SQL_ValidarTabla;
                        else Globals.Query = Properties.Resources.SQL_ValidarTabla;

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
                    if (Globals.CONF.UEXXVSIR == "1")
                    {
                        if (string.IsNullOrEmpty(ArchivoSAP)) throw new Exception("Primero seleccione el archivo SAP de compras");
                    }

                    DateTime fecha = Convert.ToDateTime(oPeriodo.Selected.Value + "-01");
                    Globals.InformationMessage("Consultando documentos, por favor espere...");
                    var documentosSIRE = ConsultaRegistroSIRE(pVal, oForm);
                    var documentosSAP = ConsultaRegistroSAP(pVal, oForm);

                    string json1 = JsonConvert.SerializeObject(documentosSIRE);
                    string json2 = JsonConvert.SerializeObject(documentosSAP);

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
                    Globals.SuccessMessage("Carga completada correctamente!");
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
                    Globals.InformationMessage("Leyendo documento, por favor espere...");
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
                            UEXXTCAMBIO = GetDouble(c, 26),
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
                    //Consulta por API SIRE


                }
                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
                    Globals.InformationMessage("Leyendo documento, por favor espere...");
                    list = File.ReadLines(Archivo).Skip(1).Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(linea =>
                    {
                        var c = linea.Split('|');

                        string ObjType = string.Empty, DocEntry = string.Empty;
                        if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ObtieneLlavesDocumentos;
                        else Globals.Query = Properties.Resources.SQL_ObtieneLlavesDocumentos;
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
                            UEXXTCAMBIO = GetDouble(c, 26),
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
                    //Cambiar por el prodimiento que desarrollará B1 agregar filtro periodo y empresa
                    Globals.Query = "CALL \"EXX_SIRE_REGCOM\"('2024')";
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
                                UEXXRUC = Globals.oRec.Fields.Item("U_EXX_RUC").Value?.ToString(),
                                UEXXRAZSOC = Globals.oRec.Fields.Item("U_EXX_RAZSOC").Value?.ToString(),
                                UEXXPERIODO = Globals.oRec.Fields.Item("U_EXX_PERIODO").Value?.ToString(),
                                UEXXCARSUN = Globals.oRec.Fields.Item("U_EXX_CARSUN").Value?.ToString(),
                                UEXXFEMI = Globals.oRec.Fields.Item("U_EXX_FEMI").Value?.ToString(),
                                UEXXFVCTO = Globals.oRec.Fields.Item("U_EXX_FVCTO").Value?.ToString(),
                                UEXXTIPDOC = Globals.oRec.Fields.Item("U_EXX_TIPDOC").Value?.ToString(),
                                UEXXSERIE = Globals.oRec.Fields.Item("U_EXX_SERIE").Value?.ToString(),
                                UEXXANIO = Globals.oRec.Fields.Item("U_EXX_ANIO").Value?.ToString(),

                                UEXXNROINI = Globals.oRec.Fields.Item("U_EXX_NROINI").Value?.ToString(),
                                UEXXNROFIN = Globals.oRec.Fields.Item("U_EXX_NROFIN").Value?.ToString(),
                                UEXXTIPIDE = Globals.oRec.Fields.Item("U_EXX_TIPIDE").Value?.ToString(),
                                UEXXNUMIDE = Globals.oRec.Fields.Item("U_EXX_NUMIDE").Value?.ToString(),
                                UEXXRAZCLI = Globals.oRec.Fields.Item("U_EXX_RAZCLI").Value?.ToString(),

                                UEXXBIGDG = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_BIGDG").Value ?? 0.00),
                                UEXXIGVDG = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_IGVDG").Value ?? 0.00),
                                UEXXBIGDGN = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_BIGDGN").Value ?? 0.00),
                                UEXXIGVDGN = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_IGVDGN").Value ?? 0.00),
                                UEXXBIDNG = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_BIDNG").Value ?? 0.00),
                                UEXXIGVDNG = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_IGVDNG").Value ?? 0.00),
                                UEXXVALNG = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_VALNG").Value ?? 0.00),
                                UEXXISC = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_ISC").Value ?? 0.00),
                                UEXXICBPER = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_ICBPER").Value ?? 0.00),
                                UEXXOTRTRI = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_OTRTRI").Value ?? 0.00),
                                UEXXTOTAL = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_TOTAL").Value ?? 0.00),

                                UEXXMONEDA = Globals.oRec.Fields.Item("U_EXX_MONEDA").Value?.ToString(),
                                UEXXTCAMBIO = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_TCAMBIO").Value ?? 0.00),
                                UEXXFEMOD = Globals.oRec.Fields.Item("U_EXX_FEMOD").Value?.ToString(),
                                UEXXTIPMOD = Globals.oRec.Fields.Item("U_EXX_TIPMOD").Value?.ToString(),
                                UEXXSERMOD = Globals.oRec.Fields.Item("U_EXX_SERMOD").Value?.ToString(),
                                UEXXCODDAM = Globals.oRec.Fields.Item("U_EXX_CODDAM").Value?.ToString(),
                                UEXXNUMMOD = Globals.oRec.Fields.Item("U_EXX_NUMMOD").Value?.ToString(),
                                UEXXCLASIF = Globals.oRec.Fields.Item("U_EXX_CLASIF").Value?.ToString(),
                                UEXXIDPROY = Globals.oRec.Fields.Item("U_EXX_IDPROY").Value?.ToString(),
                                UEXXPORPAR = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_PORPAR").Value ?? 0.0),
                                UEXXIMB = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_IMB").Value ?? 0.00),
                                UEXXCARORI = Globals.oRec.Fields.Item("U_EXX_CARORI").Value?.ToString(),
                                UEXXDETRA = Convert.ToDouble(Globals.oRec.Fields.Item("U_EXX_DETRA").Value ?? 0.00),
                                UEXXTIPNOT = Globals.oRec.Fields.Item("U_EXX_TIPNOT").Value?.ToString(),
                                UEXXESTCOM = Globals.oRec.Fields.Item("U_EXX_ESTCOM").Value?.ToString(),
                                UEXXINCAL = Globals.oRec.Fields.Item("U_EXX_INCAL").Value?.ToString(),

                                UEXXCLU1 = Globals.oRec.Fields.Item("U_EXX_CLU1").Value?.ToString(),
                                UEXXCLU2 = Globals.oRec.Fields.Item("U_EXX_CLU2").Value?.ToString(),
                                UEXXCLU3 = Globals.oRec.Fields.Item("U_EXX_CLU3").Value?.ToString(),
                                UEXXCLU4 = Globals.oRec.Fields.Item("U_EXX_CLU4").Value?.ToString(),
                                UEXXCLU5 = Globals.oRec.Fields.Item("U_EXX_CLU5").Value?.ToString(),
                                UEXXCLU6 = Globals.oRec.Fields.Item("U_EXX_CLU6").Value?.ToString(),
                                UEXXCLU7 = Globals.oRec.Fields.Item("U_EXX_CLU7").Value?.ToString(),
                                UEXXCLU8 = Globals.oRec.Fields.Item("U_EXX_CLU8").Value?.ToString(),
                                UEXXCLU9 = Globals.oRec.Fields.Item("U_EXX_CLU9").Value?.ToString(),
                                UEXXCLU10 = Globals.oRec.Fields.Item("U_EXX_CLU10").Value?.ToString(),
                                UEXXCLU11 = Globals.oRec.Fields.Item("U_EXX_CLU11").Value?.ToString(),
                                UEXXCLU12 = Globals.oRec.Fields.Item("U_EXX_CLU12").Value?.ToString(),
                                UEXXCLU13 = Globals.oRec.Fields.Item("U_EXX_CLU13").Value?.ToString(),
                                UEXXCLU14 = Globals.oRec.Fields.Item("U_EXX_CLU14").Value?.ToString(),
                                UEXXCLU15 = Globals.oRec.Fields.Item("U_EXX_CLU15").Value?.ToString(),
                                UEXXCLU16 = Globals.oRec.Fields.Item("U_EXX_CLU16").Value?.ToString(),
                                UEXXCLU17 = Globals.oRec.Fields.Item("U_EXX_CLU17").Value?.ToString(),
                                UEXXCLU18 = Globals.oRec.Fields.Item("U_EXX_CLU18").Value?.ToString(),
                                UEXXCLU19 = Globals.oRec.Fields.Item("U_EXX_CLU19").Value?.ToString(),
                                UEXXCLU20 = Globals.oRec.Fields.Item("U_EXX_CLU20").Value?.ToString(),
                                UEXXCLU21 = Globals.oRec.Fields.Item("U_EXX_CLU21").Value?.ToString(),
                                UEXXCLU22 = Globals.oRec.Fields.Item("U_EXX_CLU22").Value?.ToString(),
                                UEXXCLU23 = Globals.oRec.Fields.Item("U_EXX_CLU23").Value?.ToString(),
                                UEXXCLU24 = Globals.oRec.Fields.Item("U_EXX_CLU24").Value?.ToString(),
                                UEXXCLU25 = Globals.oRec.Fields.Item("U_EXX_CLU25").Value?.ToString(),
                                UEXXCLU26 = Globals.oRec.Fields.Item("U_EXX_CLU26").Value?.ToString(),
                                UEXXCLU27 = Globals.oRec.Fields.Item("U_EXX_CLU27").Value?.ToString(),
                                UEXXCLU28 = Globals.oRec.Fields.Item("U_EXX_CLU28").Value?.ToString(),
                                UEXXCLU29 = Globals.oRec.Fields.Item("U_EXX_CLU29").Value?.ToString(),
                                UEXXCLU30 = Globals.oRec.Fields.Item("U_EXX_CLU30").Value?.ToString(),
                                UEXXCLU31 = Globals.oRec.Fields.Item("U_EXX_CLU31").Value?.ToString(),
                                UEXXCLU32 = Globals.oRec.Fields.Item("U_EXX_CLU32").Value?.ToString(),
                                UEXXCLU33 = Globals.oRec.Fields.Item("U_EXX_CLU33").Value?.ToString(),
                                UEXXCLU34 = Globals.oRec.Fields.Item("U_EXX_CLU34").Value?.ToString(),
                                UEXXCLU35 = Globals.oRec.Fields.Item("U_EXX_CLU35").Value?.ToString(),
                                UEXXCLU36 = Globals.oRec.Fields.Item("U_EXX_CLU36").Value?.ToString(),
                                UEXXCLU37 = Globals.oRec.Fields.Item("U_EXX_CLU37").Value?.ToString(),
                                UEXXCLU38 = Globals.oRec.Fields.Item("U_EXX_CLU38").Value?.ToString(),
                                UEXXCLU39 = Globals.oRec.Fields.Item("U_EXX_CLU39").Value?.ToString()
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

            if (double.TryParse(c[index], NumberStyles.Any, CultureInfo.InvariantCulture, out valor))
                return valor;

            return null;
        }

        private static string GetDateString(string[] c, int index)
        {
            if (index >= c.Length) return string.Empty;
            if (string.IsNullOrWhiteSpace(c[index])) return string.Empty;

            return DateTime.TryParseExact(
                c[index],
                "dd/MM/yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime fecha)
                ? fecha.ToString("yyyyMMdd")
                : string.Empty;
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
