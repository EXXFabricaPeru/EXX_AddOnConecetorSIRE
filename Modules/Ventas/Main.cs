using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Modules.Ventas
{
    public class Main
    {
        public static void LoadForm(ref Form oForm)
        {
            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        internal static void FormMode(MenuEvent pVal, ref Form oForm)
        {
            throw new NotImplementedException();
        }

        internal static void FormOK( ref Form oForm)
        {
            throw new NotImplementedException();
        }
    }
}
