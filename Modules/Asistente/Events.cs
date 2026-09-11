using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Modules.Asistente
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
                            Main.ProcesarAsistente(pVal, oForm, out BubbleEvent); break;
                        case "3":
                            Main.ConsultarEmpresas(pVal, oForm, out BubbleEvent); break;
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
                if (pVal.ActionSuccess) { }
                Main.InhabilitarBoton(pVal, oForm);
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
                switch (pVal.ItemUID)
                {
                    case "mEmpresas":
                        if (pVal.ColUID == "DocEntry") Main.LinkPressedDinamic(pVal, oForm, out BubbleEvent); break;
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
