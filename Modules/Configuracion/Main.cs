using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using AddOnConectorSIRE.Utilities;
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

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("DT_0");
                SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)oForm.Items.Item("Item_1").Specific;
                oFolder.Select();

                if (Globals.IsHana())
                {
                    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxSucursal;
                    else Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxOADM;
                }
                else
                {
                    if (Globals.isMultiBranch) Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxSucursal;
                    else Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxOADM;
                }

                oDataTable.ExecuteQuery(Globals.Query);
                oMatrix.Columns.Item("BPLId").DataBind.Bind("DT_0", "BPLId");
                oMatrix.Columns.Item("GlblLocNum").DataBind.Bind("DT_0", "GlblLocNum");
                oMatrix.Columns.Item("BPLName").DataBind.Bind("DT_0", "BPLName");
                oMatrix.Columns.Item("U_EXX_APIS").DataBind.Bind("DT_0", "U_EXX_APIS");
                oMatrix.Columns.Item("U_EXX_USER").DataBind.Bind("DT_0", "U_EXX_USER");
                oMatrix.Columns.Item("U_EXX_PASS").DataBind.Bind("DT_0", "U_EXX_PASS");
                oMatrix.Columns.Item("U_EXX_CLID").DataBind.Bind("DT_0", "U_EXX_CLID");
                oMatrix.Columns.Item("BPLId").DataBind.Bind("DT_0", "BPLId");
                oMatrix.Columns.Item("U_EXX_CLSE").DataBind.Bind("DT_0", "U_EXX_CLSE");
                oMatrix.LoadFromDataSource();

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    SAPbouiCOM.EditText celda = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_EXX_PASS").Cells.Item(i).Specific);
                    celda.IsPassword = true;
                }

                oMatrix.AutoResizeColumns();
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
                    Modules.Configuracion.Main.ValidaConfig(Globals.CONF);
                }
                else
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("etSL").Specific).Value = string.Empty;
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

                //if (string.IsNullOrEmpty(empresa.UEXXAPIS) || string.IsNullOrEmpty(empresa.UEXXUSER) || string.IsNullOrEmpty(empresa.UEXXPASS) || string.IsNullOrEmpty(empresa.UEXXCLID) || string.IsNullOrEmpty(empresa.UEXXCLSE))
                //    throw new Exception("La empresa selccionada no tiene la configuración completa por favor revise en: Gestión > EXX - SIRE SUNAT > EXX - Configuración SIRE");
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
                SAPbouiCOM.Matrix oGrid = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbConSAP").Specific).Selected == null) throw new Exception("Dele seleccionar un tipo de Conexión SAP");
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSAP").Specific).Selected == null) throw new Exception("Debe seleccionar un método para la validacción de documentos SAP");
                if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSIRE").Specific).Selected == null) throw new Exception("Debe seleccionar un método para la validacción de documentos SIRE");

                UEXXSIRECONF CONFTEMP = new UEXXSIRECONF();
                CONFTEMP.Code = "01";
                CONFTEMP.UEXXCONS = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbConSAP").Specific).Selected.Value;
                CONFTEMP.UEXXVSAP = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSAP").Specific).Selected.Value;
                CONFTEMP.UEXXURSL = ((SAPbouiCOM.EditText)oForm.Items.Item("etSL").Specific).Value;
                CONFTEMP.UEXXVSIR = ((SAPbouiCOM.ComboBox)oForm.Items.Item("cbValSIRE").Specific).Selected.Value;
                CONFTEMP.APIS = new List<UEXXSIREAPIS>();

                for (int i = 1; i <= oGrid.RowCount; i++)
                {
                    CONFTEMP.APIS.Add(new UEXXSIREAPIS
                    {
                        Code = ((SAPbouiCOM.EditText)oGrid.Columns.Item("BPLId").Cells.Item(i).Specific).Value.ToString(),
                        UEXXRUC = ((SAPbouiCOM.EditText)oGrid.Columns.Item("GlblLocNum").Cells.Item(i).Specific).Value.ToString(),
                        UEXXAPIS = ((SAPbouiCOM.EditText)oGrid.Columns.Item("U_EXX_APIS").Cells.Item(i).Specific).Value.ToString(),
                        UEXXUSER = ((SAPbouiCOM.EditText)oGrid.Columns.Item("U_EXX_USER").Cells.Item(i).Specific).Value.ToString(),
                        UEXXPASS = ((SAPbouiCOM.EditText)oGrid.Columns.Item("U_EXX_PASS").Cells.Item(i).Specific).Value.ToString(),
                        UEXXCLID = ((SAPbouiCOM.EditText)oGrid.Columns.Item("U_EXX_CLID").Cells.Item(i).Specific).Value.ToString(),
                        UEXXCLSE = ((SAPbouiCOM.EditText)oGrid.Columns.Item("U_EXX_CLSE").Cells.Item(i).Specific).Value.ToString()
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
                Globals.existeConf = true;
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

        public static void SetearPassword(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string Password = ((SAPbouiCOM.EditText)oForm.Items.Item("etPass").Specific).Value.Trim();
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (oMatrix.IsRowSelected(i))
                    {
                        if (string.IsNullOrEmpty(Password))
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_EXX_PASS").Cells.Item(i).Specific).Value = Password;
                        else
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_EXX_PASS").Cells.Item(i).Specific).Value = ExxisEncryptor.Encrypt(Password);

                        return;
                    }
                }
                Globals.ErrorMessage("Debe seleccionar una fila de la lista para actualizar el password.");
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
                            oForm.Items.Item("mEmpresas").Visible = false;
                            oForm.Items.Item("stPass").Visible = false;
                            oForm.Items.Item("etPass").Visible = false;
                            oForm.Items.Item("3").Visible = false;
                        }
                        else
                        {
                            oForm.Items.Item("mEmpresas").Visible = true;
                            oForm.Items.Item("stPass").Visible = true;
                            oForm.Items.Item("etPass").Visible = true;
                            oForm.Items.Item("3").Visible = true;

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mEmpresas").Specific;
                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                SAPbouiCOM.EditText celda = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_EXX_PASS").Cells.Item(i).Specific);
                                celda.IsPassword = true;
                            }
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
