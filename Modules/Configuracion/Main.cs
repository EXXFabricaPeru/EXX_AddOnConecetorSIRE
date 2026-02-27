using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Form = SAPbouiCOM.Form;

namespace AddOnConectorSIRE.Modules.Configuracion
{
    public class Main
    {
        public static void LoadForm()
        {
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            try
            {
                oForm = Globals.SBO_Application.Forms.Item("EXX_SIRE_CONF");
                Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto.");
            }
            catch
            {
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = (SAPbouiCOM.FormCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "EXX_SIRE_CONF";
                fcp.UniqueID = "EXX_SIRE_CONF";
                string FormName = "\\SetupFields\\Forms\\SRF_EXX_SIRE_CONF.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);
                int centerX = (Globals.SBO_Application.Desktop.Width - oForm.Width) / 2;
                int centerY = ((Globals.SBO_Application.Desktop.Height - Convert.ToInt32(Globals.SBO_Application.Desktop.Height * 0.15)) - oForm.Height) / 2;
                oForm.Left = centerX;
                oForm.Top = centerY;

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("gEmpresas").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");
                SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)oForm.Items.Item("Item_1").Specific;
                oFolder.Select();

                if (Globals.IsHana())
                {
                    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxSucursal;
                    else Globals.Query = Properties.Resources.HANA_MultiBranch;
                }
                else
                {
                    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxSucursal;
                    else Globals.Query = Properties.Resources.SQL_MultiBranch;
                }
                oDataTable.ExecuteQuery(Globals.Query);
                oGrid.Columns.Item("BPLId").Editable = false;
                oGrid.Columns.Item("BPLId").TitleObject.Caption = "Sucursal";
                oGrid.Columns.Item("GlblLocNum").Editable = false;
                oGrid.Columns.Item("GlblLocNum").TitleObject.Caption = "RUC";
                oGrid.Columns.Item("BPLName").Editable = false;
                oGrid.Columns.Item("BPLName").TitleObject.Caption = "Nombre Empresa";
                oGrid.Columns.Item("U_EXX_APIS").Editable = true;
                oGrid.Columns.Item("U_EXX_APIS").TitleObject.Caption = "URL API";
                oGrid.Columns.Item("U_EXX_USER").Editable = true;
                oGrid.Columns.Item("U_EXX_USER").TitleObject.Caption = "Usuario SOL";
                oGrid.Columns.Item("U_EXX_PASS").Editable = true;
                oGrid.Columns.Item("U_EXX_PASS").TitleObject.Caption = "Clave SOL";
                oGrid.Columns.Item("U_EXX_CLID").Editable = true;
                oGrid.Columns.Item("U_EXX_CLID").TitleObject.Caption = "Client ID";
                oGrid.Columns.Item("U_EXX_CLSE").Editable = true;
                oGrid.Columns.Item("U_EXX_CLSE").TitleObject.Caption = "Client Secret";
                oGrid.AutoResizeColumns();
                if (Globals.existeConf)
                {
                    if (Globals.IsHana())
                        Globals.Query = Properties.Resources.HANA_ObtenerConfiguracion;
                    else
                        Globals.Query = Properties.Resources.SQL_ObtenerConfiguracion;
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.RecordCount == 0)
                        Globals.existeConf = false;

                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbConSAP").Specific).Select(Globals.CONF.UEXXCONS, BoSearchKey.psk_ByValue);
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSAP").Specific).Select(Globals.CONF.UEXXVSAP, BoSearchKey.psk_ByValue);
                    ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSIRE").Specific).Select(Globals.CONF.UEXXVSIR, BoSearchKey.psk_ByValue);
                    ((SAPbouiCOM.EditText)oForm.Items.Item("etSL").Specific).Value = Globals.CONF.UEXXURSL;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etAPI").Specific).Value = Globals.CONF.UEXXAPIS;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etUser").Specific).Value = Globals.CONF.UEXXUSER;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etPass").Specific).Value = Globals.CONF.UEXXPASS;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etClient").Specific).Value = Globals.CONF.UEXXCLID;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etSecret").Specific).Value = Globals.CONF.UEXXCLSE;
                    Modules.Configuracion.Main.ValidaConfig(Globals.CONF);
                }
                else
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("etSL").Specific).Value = string.Empty;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etAPI").Specific).Value = string.Empty;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etUser").Specific).Value = string.Empty;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etPass").Specific).Value = string.Empty;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etClient").Specific).Value = string.Empty;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("etSecret").Specific).Value = string.Empty;
                }
            }
            oForm.Refresh();
            oForm.Visible = true;
        }

        public static bool ValidaConfig(UEXXSIRECONF CONF)
        {
            try
            {
                if (string.IsNullOrEmpty(CONF.UEXXCONS)) throw new Exception("Dele seleccionar un tipo de Conexión SAP");
                else if (CONF.UEXXCONS == "2" && string.IsNullOrEmpty(CONF.UEXXURSL)) throw new Exception("Para el tipo de conexión SAP 'Service layer' debe llenar el campo URL Service Layer");
                if (string.IsNullOrEmpty(CONF.UEXXVSAP)) throw new Exception("Debe seleccionar un método para la validacción de documentos SAP");
                if (string.IsNullOrEmpty(CONF.UEXXVSIR)) throw new Exception("Debe seleccionar un método para la validacción de documentos SIRE");
                //else if (CONF.UEXXVSIR == "2" && string.IsNullOrEmpty(CONF.UEXXAPIS)) throw new Exception("Para el tipo de validación SIRE 'API SIRE' debe llenar el campo URL API SIRE");
                //else if (CONF.UEXXVSIR == "2" && string.IsNullOrEmpty(CONF.UEXXUSER)) throw new Exception("Para el tipo de validación SIRE 'API SIRE' debe llenar el campo Usuario SOL");
                //else if (CONF.UEXXVSIR == "2" && string.IsNullOrEmpty(CONF.UEXXPASS)) throw new Exception("Para el tipo de validación SIRE 'API SIRE' debe llenar el campo Clave SOL");
                //else if (CONF.UEXXVSIR == "2" && string.IsNullOrEmpty(CONF.UEXXCLID)) throw new Exception("Para el tipo de validación SIRE 'API SIRE' debe llenar el campo Client-Id");
                //else if (CONF.UEXXVSIR == "2" && string.IsNullOrEmpty(CONF.UEXXCLSE)) throw new Exception("Para el tipo de validación SIRE 'API SIRE' debe llenar el campo Client-Secret");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        public static void GuardarConfig(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("gEmpresas").Specific;
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbConSAP").Specific).Selected == null) throw new Exception("Dele seleccionar un tipo de Conexión SAP");
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSAP").Specific).Selected == null) throw new Exception("Debe seleccionar un método para la validacción de documentos SAP");
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSIRE").Specific).Selected == null) throw new Exception("Debe seleccionar un método para la validacción de documentos SIRE");

                UEXXSIRECONF CONFTEMP = new UEXXSIRECONF();
                CONFTEMP.Code = "01";
                CONFTEMP.UEXXCONS = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbConSAP").Specific).Selected.Value;
                CONFTEMP.UEXXVSAP = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSAP").Specific).Selected.Value;
                CONFTEMP.UEXXURSL = ((SAPbouiCOM.EditText)oForm.Items.Item("etSL").Specific).Value;
                CONFTEMP.UEXXVSIR = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSIRE").Specific).Selected.Value;
                //CONFTEMP.UEXXAPIS = ((SAPbouiCOM.EditText)oForm.Items.Item("etAPI").Specific).Value;
                //CONFTEMP.UEXXUSER = ((SAPbouiCOM.EditText)oForm.Items.Item("etUser").Specific).Value;
                //CONFTEMP.UEXXPASS = ((SAPbouiCOM.EditText)oForm.Items.Item("etPass").Specific).Value;
                //CONFTEMP.UEXXCLID = ((SAPbouiCOM.EditText)oForm.Items.Item("etClient").Specific).Value;
                //CONFTEMP.UEXXCLSE = ((SAPbouiCOM.EditText)oForm.Items.Item("etSecret").Specific).Value;
                CONFTEMP.APIS = new List<UEXXSIREAPIS>();

                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    CONFTEMP.APIS.Add(new UEXXSIREAPIS
                    {
                        Code = oGrid.DataTable.GetValue("BPLId", i).ToString(),
                        UEXXRUC = oGrid.DataTable.GetValue("GlblLocNum", i).ToString(),
                        UEXXAPIS = oGrid.DataTable.GetValue("U_EXX_APIS", i).ToString(),
                        UEXXUSER = oGrid.DataTable.GetValue("U_EXX_USER", i).ToString(),
                        UEXXPASS = oGrid.DataTable.GetValue("U_EXX_PASS", i).ToString(),
                        UEXXCLID = oGrid.DataTable.GetValue("U_EXX_CLID", i).ToString(),
                        UEXXCLSE = oGrid.DataTable.GetValue("U_EXX_CLSE", i).ToString()
                    });
                }

                ValidaConfig(CONFTEMP);
                Globals.StartTransaction();
                if (Globals.existeConf)
                    SAPSDK.ActualizarSireConf(CONFTEMP);
                else
                    SAPSDK.GuardarSireConf(CONFTEMP);

                foreach (var empresa in CONFTEMP.APIS)
                {
                    if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ExisteConfAPI;
                    else Globals.Query = Properties.Resources.SQL_ExisteConfAPI;
                    Globals.Query = string.Format(Globals.Query, empresa.Code);
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.RecordCount == 0) SAPSDK.CrearRegistroAPI(empresa);
                    else SAPSDK.ActualizarRegistroAPI(empresa);
                    Globals.Release(Globals.oRec);
                }
                Globals.CommitTransaction();

                Globals.CONF = CONFTEMP;
                Setup.ValidarVersion();
                Menu.LoadMenu();
                Globals.SuccessMessage("Configuración guardada satisfactoriamente.");
            }
            catch (Exception ex)
            {
                if (Globals.InTransaction()) Globals.RollBackTransaction();
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void CerrarFormulario(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!Globals.existeConf)
                {
                    int seleccion = Globals.SBO_Application.MessageBox($"Si no guarda una configuración el {Globals.AddOnName} no podrá iniciar los módulos principales. ¿Continuar cerrando?", 1, "Si", "No");
                    if (seleccion != 1)
                        BubbleEvent = false;
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void MostrosOcultarCampos(ItemEvent pVal, Form oForm)
        {
            try
            {
                string value = string.Empty;
                value = ((SAPbouiCOM.ComboBox)oForm.Items.Item(pVal.ItemUID).Specific).Selected.Value;
                switch (pVal.ItemUID)
                {
                    case "cbConSAP":
                        if (value == "1")
                        {
                            oForm.Items.Item("stSL").Visible = false;
                            oForm.Items.Item("etSL").Visible = false;
                        }
                        else
                        {
                            oForm.Items.Item("stSL").Visible = true;
                            oForm.Items.Item("etSL").Visible = true;
                        }
                        break;
                    case "cbValSIRE":
                        if (value == "1")
                        {
                            oForm.Items.Item("gEmpresas").Visible = false;
                            //oForm.Items.Item("stAPI").Visible = false;
                            //oForm.Items.Item("etAPI").Visible = false;
                            //oForm.Items.Item("stUser").Visible = false;
                            //oForm.Items.Item("etUser").Visible = false;
                            //oForm.Items.Item("stPass").Visible = false;
                            //oForm.Items.Item("etPass").Visible = false;
                            //oForm.Items.Item("stClient").Visible = false;
                            //oForm.Items.Item("etClient").Visible = false;
                            //oForm.Items.Item("stSecret").Visible = false;
                            //oForm.Items.Item("etSecret").Visible = false;
                        }
                        else
                        {
                            oForm.Items.Item("gEmpresas").Visible = true;
                            //oForm.Items.Item("stAPI").Visible = true;
                            //oForm.Items.Item("etAPI").Visible = true;
                            //oForm.Items.Item("stUser").Visible = true;
                            //oForm.Items.Item("etUser").Visible = true;
                            //oForm.Items.Item("stPass").Visible = true;
                            //oForm.Items.Item("etPass").Visible = true;
                            //oForm.Items.Item("stClient").Visible = true;
                            //oForm.Items.Item("etClient").Visible = true;
                            //oForm.Items.Item("stSecret").Visible = true;
                            //oForm.Items.Item("etSecret").Visible = true;
                        }
                        break;
                }
                oForm.Refresh();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
