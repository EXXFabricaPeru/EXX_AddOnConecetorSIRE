using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE
{
    public class Menu
    {
        public static void LoadMenu()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage = (SAPbouiCOM.MenuCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            try
            {
                #region Producción de Flejes
                oMenuItem = Globals.SBO_Application.Menus.Item("43520");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "EXX_SIRE";
                oCreationPackage.String = "EXX - SIRE SUNAT";
                oCreationPackage.Position = oMenus.Count;
                #region ExisteMenu
                try
                {
                    if (oMenus.Exists("EXX_SIRE"))
                    {
                        Globals.SBO_Application.Menus.RemoveEx("EXX_SIRE");
                    }
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception)
                { }
                oMenuItem = Globals.SBO_Application.Menus.Item("EXX_SIRE");
                oMenus = oMenuItem.SubMenus;
                #endregion
                #endregion

                #region Configuración SIRE
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EXX_SIRE_CONF";
                oCreationPackage.String = "EXX - Configuración SIRE";
                #region ExisteMenu
                try
                {
                    if (oMenus.Exists("EXX_SIRE_CONF"))
                    {
                        Globals.SBO_Application.Menus.RemoveEx("EXX_SIRE_CONF");
                    }
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception)
                { }
                #endregion
                #endregion

                if (Globals.existeConf)
                {
                    #region Registro de Compras
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "EXX_SIRE_COMP";
                    oCreationPackage.String = "EXX - Registro de Compras";
                    #region ExisteMenu
                    try
                    {
                        if (oMenus.Exists("EXX_SIRE_COMP"))
                        {
                            Globals.SBO_Application.Menus.RemoveEx("EXX_SIRE_COMP");
                        }
                        oMenus.AddEx(oCreationPackage);
                    }
                    catch (Exception)
                    { }
                    #endregion
                    #endregion

                    #region Registro de Ventas
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "EXX_SIRE_VENT";
                    oCreationPackage.String = "EXX - Registro de Ventas";
                    #region ExisteMenu
                    try
                    {
                        if (oMenus.Exists("EXX_SIRE_VENT"))
                        {
                            Globals.SBO_Application.Menus.RemoveEx("EXX_SIRE_VENT");
                        }
                        oMenus.AddEx(oCreationPackage);
                    }
                    catch (Exception)
                    { }
                    #endregion
                    #endregion
                }
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }
    }
}
