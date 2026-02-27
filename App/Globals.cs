using AddOnConectorSIRE.Entities;
using Newtonsoft.Json.Linq;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using Application = SAPbouiCOM.Framework.Application;

namespace AddOnConectorSIRE
{
    public class Globals
    {
        public static RevalMain Addon;
        public static String ShortName = "(EXX)";
        public static string AddOnName = "AddOn Conector SIRE";
        public static string AddOnVersion = "1.0.0.0";
        public static int continuar = -1;
        public static string Query = null;
        public static SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
        public static SAPbouiCOM.Application SBO_Application = Application.SBO_Application;
        public static SAPbobsCOM.Company oCompany;
        public static int sErrCode;
        public static string sErrMsg = null;

        public static Application oApp;

        public const string LinkedSystemObject = "1";
        public const string LinkedUDO = "2";
        public const string LinkedTable = "3";

        public static bool OnlyURI = false;
        public static List<string> CamposNumericos = new List<string> { "FolioNumber", "DocEntry", "DocNum", "Invoices/DocEntry", "Invoices/DocumentLines/BaseEntry", "Series" };
        public static bool existeConf = false;
        public static UEXXSIRECONF CONF;
        public static bool MultiEmpresa = false;
        public static string B1SESSION = string.Empty;
        public static string configPath = AppDomain.CurrentDomain.BaseDirectory + "/SetupFields/";
        public static bool isMultiBranch = false;

        public static class Operador
        {
            public const string Igual = " eq ";
            public const string Diferente = " ne ";
            public const string Mayor = " rq ";
            public const string Menor = " lq ";
            public const string MayorIgual = " ge ";
            public const string MenorIgual = " le ";
        }

        public static List<T> ReadJson<T>(string FileName) where T : class, new()
        {
            try
            {
                if (System.IO.File.Exists(configPath + FileName))
                {
                    string jsonContent = System.IO.File.ReadAllText(configPath + FileName);
                    JArray contentArray = JArray.Parse(jsonContent);
                    return contentArray.ToObject<List<T>>();
                }
                else
                    throw new Exception($"Archivo {FileName} no encontrado.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static SAPbobsCOM.Recordset RunQuery(string query)
        {
            try
            {
                oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(query);

                return oRec;
            }
            catch (Exception ex)
            {
                MessageBox(ex.Message);
                return null;
            }
        }

        public static void StartTransaction()
        {
            try
            {
                oCompany.StartTransaction();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static bool InTransaction()
        {
            try
            {
                return oCompany.InTransaction;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void CommitTransaction()
        {
            try
            {
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void RollBackTransaction()
        {
            try
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static object Release(object objeto)
        {
            if (objeto != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
                Query = null;
                GC.Collect();
            }
            return null;
        }

        public static void ErrorMessage(String msg)
        {
            SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short);
        }

        public static void InformationMessage(String msg)
        {
            SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public static void SuccessMessage(String msg)
        {
            SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }
        public static void SuccessMessage2(String msg)
        {
            SBO_Application.StatusBar.SetText(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        public static void MessageBox(String msg)
        {
            SBO_Application.MessageBox(msg);
        }

        public static void LoadForm(String Frm)
        {
            SAPbouiCOM.MenuItem menu = SBO_Application.Menus.Item("47616");
            try
            {
                if (menu.SubMenus.Count > 0)
                {
                    for (int i = 0; i < menu.SubMenus.Count; i++)
                    {
                        if (menu.SubMenus.Item(i).String.Contains(Frm))
                        {
                            menu.SubMenus.Item(i).Activate();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("ERROR: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, true);
            }
        }

        public static void LlenarCombo(SAPbouiCOM.Form oForm, string ncombo, string qcombo, bool optLinea, string valLinea)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(ncombo).Specific;
            Globals.RunQuery(qcombo);
            Globals.oRec.MoveFirst();

            while (oCombo.ValidValues.Count > 0)
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);

            if (optLinea) { oCombo.ValidValues.Add(valLinea, valLinea); }
            while (!Globals.oRec.EoF)
            {
                oCombo.ValidValues.Add(Globals.oRec.Fields.Item(0).Value.ToString(), Globals.oRec.Fields.Item(1).Value.ToString());
                Globals.oRec.MoveNext();
            }
            Globals.Release(oCombo);
            Globals.Release(Globals.oRec);
        }

        public static System.Data.DataTable RecorsetToDataTable(SAPbobsCOM.Recordset SAPRecordset)
        {
            System.Data.DataTable dtTable = new System.Data.DataTable();
            System.Data.DataColumn NewCol;
            DataRow NewRow;
            int ColCount;

            try
            {
                for (ColCount = 0; ColCount < SAPRecordset.Fields.Count; ColCount++)
                {
                    NewCol = new System.Data.DataColumn(SAPRecordset.Fields.Item(ColCount).Name);
                    dtTable.Columns.Add(NewCol);
                }

                while (!SAPRecordset.EoF)
                {
                    NewRow = dtTable.NewRow();
                    for (ColCount = 0; ColCount < SAPRecordset.Fields.Count; ColCount++)
                        NewRow[SAPRecordset.Fields.Item(ColCount).Name] = SAPRecordset.Fields.Item(ColCount).Value;

                    dtTable.Rows.Add(NewRow);

                    SAPRecordset.MoveNext();
                }

                return dtTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string LoadFromXML(ref string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;
            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath;
            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }

        public static string BuildURI<T>(T @object)
        {
            try
            {
                var type = typeof(T);
                string uri = string.Empty;
                var entityAttribute = (SAPEntityNameAttribute)Attribute.GetCustomAttribute(type, typeof(SAPEntityNameAttribute));
                if (entityAttribute == null)
                {
                    throw new Exception("El atributo SAPEntityName no está definido en la clase.");
                }

                var sapKeyProperty = type.GetProperties().FirstOrDefault(prop => Attribute.IsDefined(prop, typeof(SAPKeyAttribute)));
                if (sapKeyProperty == null)
                    return entityAttribute.Uri;

                if (Globals.OnlyURI)
                    return entityAttribute.Uri;

                var sapKeyValue = sapKeyProperty.GetValue(@object)?.ToString();
                if (string.IsNullOrEmpty(sapKeyValue))
                    return entityAttribute.Uri;
                else
                {
                    sapKeyValue = IsNumeric(sapKeyValue) ? sapKeyValue : $"'{sapKeyValue}'";
                    if (entityAttribute.Uri == "SQLQueries")
                        return $"{entityAttribute.Uri}({sapKeyValue})/List";
                    else
                        return $"{entityAttribute.Uri}({sapKeyValue})";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.OnlyURI = false;
            }
        }

        public static bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        internal static string BuildFilter(Tuple<string, string, string>[] prms)
        {
            try
            {
                string uri = string.Empty;
                if (prms != null && prms.Count() > 0)
                {
                    uri += string.Join(" and ", prms.ToList().Select(f =>
                        f.Item3 == null ? $"{f.Item1} {f.Item2} null"
                        : Globals.CamposNumericos.Any(x => x == f.Item1) ? $"{f.Item1} {f.Item2} {f.Item3}"
                        : IsNumeric(f.Item3)
                        ? $"{f.Item1} {f.Item2} '{f.Item3}'"
                        : $"{f.Item1} {f.Item2} '{f.Item3}'"
                    ));
                }

                return uri;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool IsHana()
        {
            try
            {
                if (Globals.oCompany.DbServerType == (SAPbobsCOM.BoDataServerTypes)9)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return false;
            }
        }
    }
}
