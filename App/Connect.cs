using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOnConectorSIRE
{
    public class Connect
    {
        public static void SetApplication()
        {
            try
            {
                string args = (string)Environment.GetCommandLineArgs().GetValue(1);
                if (args.Length < 1)
                    Globals.oApp = new SAPbouiCOM.Framework.Application();
                else
                    Globals.oApp = new SAPbouiCOM.Framework.Application(args);

                Globals.SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ConnectToCompany()
        {
            try
            {
                SAPbobsCOM.Company oCompany;
                oCompany = (SAPbobsCOM.Company)Globals.SBO_Application.Company.GetDICompany();
                Globals.oCompany = oCompany;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
