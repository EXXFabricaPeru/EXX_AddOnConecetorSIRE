using SAPbouiCOM;
using System;

namespace AddOnConectorSIRE.Modules.Ventas
{
    public class Events
    {
        public static void ChooseFromList(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void ItemPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            Main.ValidarRegistro(pVal, oForm, out BubbleEvent); break;
                        case "3":
                        case "4":
                            Main.SeleccionarArchivo(pVal, oForm, out BubbleEvent); break;
                        case "5":
                            Main.ProcesarRegistroVenta(pVal, oForm, out BubbleEvent); break;
                        case "6":
                            Main.ReemplazarPropuesta(pVal, oForm, out BubbleEvent); break;
                        case "7":
                            Main.ConsultaTicket(pVal, oForm, out BubbleEvent); break;
                        case "0_U_G":
                            Main.SeleccionarFila(pVal, oForm, out BubbleEvent); break;
                    }
                }

                if (pVal.ActionSuccess)
                {
                    switch (pVal.ItemUID)
                    {
                        case "1":
                            if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                Globals.SBO_Application.ActivateMenuItem("1289");
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

        public static void MatrixLinkedPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.ItemUID)
                    {
                        case "0_U_G":
                            if (pVal.ColUID == "C_0_2") Main.LinkPressedDinamic(pVal, oForm, out BubbleEvent); break;
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
    }
}
