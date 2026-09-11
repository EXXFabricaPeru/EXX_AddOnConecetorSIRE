using SAPbouiCOM;
using System;

namespace AddOnConectorSIRE.Modules.Configuracion
{
    public class Events
    {
        public static void ItemPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.ItemUID)
                    {
                        case "OK":
                            Main.GuardarConfig(pVal, oForm, out BubbleEvent); break;
                        case "2":
                            Main.CerrarFormulario(pVal, oForm, out BubbleEvent); break;
                        case "3":
                            Main.SetearPassword(pVal, oForm, out BubbleEvent); break;
                    }
                }

                if (pVal.ActionSuccess)
                {
                    switch (pVal.ItemUID)
                    {
                        case "OK":
                            oForm.Close();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void ComboSelect(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                    Main.MostrosOcultarCampos(pVal, oForm);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void FormClose(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                    Main.CerrarFormulario(pVal, oForm, out BubbleEvent); 
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }
    }
}
