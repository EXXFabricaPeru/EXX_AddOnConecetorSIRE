using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }

    class ClsFolderFileDialog
    {


        public string ruta;

        public string FindFile()
        {
            try
            {
                var variable_temp = string.Empty;
                var explorer = new FileExplorer();
                explorer.ShowFolderBrowser(false, "TXT File (*.txt)|*.txt");

                if (explorer.Error)
                {
                    //ShowMessage(explorer.LastException.Message);
                }
                else
                {
                    if (explorer.Files.Length > 0)
                    {
                        variable_temp = explorer.Files[0];

                    }
                }

                if (!string.IsNullOrEmpty(variable_temp))
                {
                    return variable_temp;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return string.Empty;
        }
    }
}
