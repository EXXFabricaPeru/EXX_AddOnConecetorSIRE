using SAPbouiCOM;
using System;

namespace AddOnConectorSIRE
{
    public class RevalMain
    {
        public RevalMain()
        {
            try
            {
                Connect.SetApplication();
                Connect.ConnectToCompany();
                Globals.SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Menu.LoadMenu();
                Setup.CargarConfiguracion();
                Globals.SetDecimalSeparator();

                if (Globals.existeConf)
                    Setup.ValidarVersion();

                Menu.LoadMenu();
                Globals.SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                Globals.SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                Globals.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                Globals.SBO_Application.StatusBar.SetText(Globals.AddOnName + " v." + Globals.AddOnVersion + " Conectada con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, (SAPbouiCOM.BoStatusBarMessageType)SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                Globals.oApp.Run();
            }
            catch (Exception ex)
            {
                Globals.ErrorMessage(ex.Message);
            }
        }

        #region Events Modules
        private void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm;
                if (!pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "EXX_SIRE_CONF":
                            Modules.Configuracion.Main.LoadForm();
                            break;
                        case "EXX_SIRE_COMP":
                            Globals.LoadForm("EXX_SIRE_COMP");
                            oForm = Globals.SBO_Application.Forms.ActiveForm;
                            Modules.Compras.Main.LoadForm(ref oForm);
                            break;
                        case "EXX_SIRE_VENT":
                            Globals.LoadForm("EXX_SIRE_VENT");
                            oForm = Globals.SBO_Application.Forms.ActiveForm;
                            Modules.Ventas.Main.LoadForm(ref oForm);
                            break;
                        case "EXX_SIRE_ASIS":
                            Modules.Asistente.Main.LoadForm();
                            break;
                        case "1281": //Buscar
                        case "1282": //Nuevo
                            oForm = Globals.SBO_Application.Forms.ActiveForm;
                            switch (oForm.TypeEx)
                            {
                                case "UDO_FT_EXX_SIRE_COMP":
                                    Modules.Compras.Main.FormMode(pVal, ref oForm);
                                    break;
                                case "UDO_FT_EXX_SIRE_VENT":
                                    Modules.Ventas.Main.FormMode(pVal, ref oForm);
                                    break;
                            }
                            break;
                        case "1304": //Actualizar
                            oForm = Globals.SBO_Application.Forms.ActiveForm;
                            switch (oForm.TypeEx)
                            {
                                case "UDO_FT_EXX_SIRE_COMP":
                                    Modules.Compras.Main.FormOK(ref oForm);
                                    break;
                                case "UDO_FT_EXX_SIRE_VENT":
                                    Modules.Ventas.Main.FormOK(ref oForm);
                                    break;
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.ErrorMessage(ex.Message);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx != "0")
            {
                try
                {
                    SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(pVal.FormUID);
                    switch (pVal.FormTypeEx)
                    {
                        case "EXX_SIRE_CONF":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ITEM_PRESSED:
                                    Modules.Configuracion.Events.ItemPressed(ref pVal, oForm, out BubbleEvent); break;
                                case BoEventTypes.et_COMBO_SELECT:
                                    Modules.Configuracion.Events.ComboSelect(ref pVal, oForm, out BubbleEvent); break;
                                case BoEventTypes.et_FORM_CLOSE:
                                    Modules.Configuracion.Events.FormClose(ref pVal, oForm, out BubbleEvent); break;
                            }
                            break;
                        case "UDO_FT_EXX_SIRE_COMP":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ITEM_PRESSED:
                                    Modules.Compras.Events.ItemPressed(ref pVal, oForm, out BubbleEvent);
                                    break;
                                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    Modules.Compras.Events.MatrixLinkedPressed(ref pVal, oForm, out BubbleEvent);
                                    break;
                            }
                            break;
                        case "UDO_FT_EXX_SIRE_VENT":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ITEM_PRESSED:
                                    Modules.Ventas.Events.ItemPressed(ref pVal, oForm, out BubbleEvent);
                                    break;
                                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    Modules.Ventas.Events.MatrixLinkedPressed(ref pVal, oForm, out BubbleEvent);
                                    break;
                            }
                            break;
                        case "EXX_SIRE_ASIS":
                            switch (pVal.EventType)
                            {
                                case BoEventTypes.et_ITEM_PRESSED:
                                    Modules.Asistente.Events.ItemPressed(ref pVal, oForm, out BubbleEvent); break;
                                case BoEventTypes.et_COMBO_SELECT:
                                    Modules.Asistente.Events.ComboSelect(ref pVal, oForm, out BubbleEvent); break;
                                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                                    Modules.Asistente.Events.MatrixLinkedPressed(ref pVal, oForm, out BubbleEvent); break;
                            }
                            break;
                    }
                }
                catch (Exception ex)
                {
                    BubbleEvent = false;
                    if (ex.Message != "Form - Invalid Form" && ex.Message != "Invalid Choose From List  [66000-104]" && !ex.Message.Contains("focus"))
                        Globals.ErrorMessage(ex.Message);
                }
            }
        }

        private void SBO_Application_FormDataEvent(ref BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(pVal.FormUID);
                if (pVal.FormTypeEx != "0")
                {
                    if (pVal.ActionSuccess)
                    {
                        if (!oForm.Title.ToUpper().Contains("CANCEL"))
                        {
                            switch (pVal.FormTypeEx)
                            {
                                case "UDO_FT_EXX_SIRE_COMP":
                                    switch (pVal.EventType)
                                    {
                                        case BoEventTypes.et_FORM_DATA_LOAD:
                                        case BoEventTypes.et_FORM_DATA_UPDATE:
                                            Modules.Compras.Main.FormOK(ref oForm);
                                            break;
                                    }
                                    break;
                                case "UDO_FT_EXX_SIRE_VENT":
                                    switch (pVal.EventType)
                                    {
                                        case BoEventTypes.et_FORM_DATA_LOAD:
                                        case BoEventTypes.et_FORM_DATA_UPDATE:
                                            Modules.Ventas.Main.FormOK(ref oForm);
                                            break;
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                if (!ex.Message.Contains("7780"))
                    if (ex.Message != "Form - Invalid Form")
                        Globals.ErrorMessage(ex.Message);
            }
        }
        #endregion

        private void SBO_Application_AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_FontChanged:
                case BoAppEventTypes.aet_LanguageChanged:
                    Globals.SBO_Application.MessageBox(Globals.AddOnName + " finalizará");
                    Environment.Exit(Environment.ExitCode);
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_ServerTerminition:
                case BoAppEventTypes.aet_ShutDown:
                    Environment.Exit(Environment.ExitCode);
                    System.Windows.Forms.Application.Exit();
                    break;
                default:
                    break;
            }
        }
    }
}
