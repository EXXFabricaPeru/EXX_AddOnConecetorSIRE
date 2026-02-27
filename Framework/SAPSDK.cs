using AddOnConectorSIRE.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace AddOnConectorSIRE.Framework
{
    public class SAPSDK
    {
        public static bool CrearTabla(UserTablesMd table)
        {
            SAPbobsCOM.UserTablesMD oTablaUser = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                if (!oTablaUser.GetByKey(table.TableName))
                {
                    oTablaUser.TableName = table.TableName;
                    oTablaUser.TableDescription = table.TableDescription;

                    SAPbobsCOM.BoUTBTableType tipo;
                    if (Enum.TryParse(table.TableType, out tipo)) oTablaUser.TableType = tipo;
                    else throw new Exception("Valor inválido para BoUTBTableType");

                    int RetVal = oTablaUser.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oTablaUser);
            }
        }

        public static bool CrearCampo(UserFieldsMd field)
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string cadena = "select \"FieldID\" from CUFD where (\"TableID\"='" + field.TableName + "' or \"TableID\"='@" + field.TableName + "') and \"AliasID\"='" + field.Name + "'";
            rs.DoQuery(cadena);

            existeCampo = rs.RecordCount;
            int FieldID = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;

            SAPbobsCOM.UserFieldsMD oCampo = (SAPbobsCOM.UserFieldsMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (existeCampo == 0)
                {
                    oCampo.TableName = field.TableName;
                    oCampo.Name = field.Name;
                    oCampo.Description = field.Description;

                    SAPbobsCOM.BoFieldTypes Type;
                    if (Enum.TryParse(field.Type, out Type)) oCampo.Type = Type;
                    else throw new Exception("Valor inválido para BoFieldTypes");

                    if (!string.IsNullOrEmpty(field.SubType))
                    {
                        SAPbobsCOM.BoFldSubTypes SubType;
                        if (Enum.TryParse(field.SubType, out SubType)) oCampo.SubType = SubType;
                        else throw new Exception("Valor inválido para BoFldSubTypes");
                    }

                    oCampo.Mandatory = string.IsNullOrEmpty(field.Mandatory) ? SAPbobsCOM.BoYesNoEnum.tNO : field.Mandatory == "tYES" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    if ((field.EditSize ?? 0) > 0)
                    {
                        oCampo.EditSize = field.EditSize ?? 0;
                    }

                    if (!string.IsNullOrEmpty(field.LinkedSystemObject))
                    {
                        SAPbobsCOM.UDFLinkedSystemObjectTypesEnum TableSystemObject;
                        if (Enum.TryParse(field.LinkedSystemObject, out TableSystemObject))
                            oCampo.LinkedSystemObject = TableSystemObject;
                        else
                            throw new Exception("Valor inválido para UDFLinkedSystemObjectTypesEnum");
                    }

                    if (!string.IsNullOrEmpty(field.LinkedTable))
                        oCampo.LinkedTable = field.LinkedTable;

                    if (!string.IsNullOrEmpty(field.LinkedUDO))
                        oCampo.LinkedUDO = field.LinkedUDO;

                    if (field.ValidValuesMD != null)
                    {
                        foreach (ValidValuesMd ValidValue in field.ValidValuesMD)
                        {
                            oCampo.ValidValues.Value = ValidValue.Value;
                            oCampo.ValidValues.Description = ValidValue.Description;
                            oCampo.ValidValues.Add();
                        }
                    }

                    if (!string.IsNullOrEmpty(field.DefaultValue))
                        oCampo.DefaultValue = field.DefaultValue;

                    int RetVal = oCampo.Add();
                    if (RetVal != 0)
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oCampo);
            }
        }

        public static bool CrearUDO(UserObjectsMd udo)
        {
            try
            {


                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void GuardarSireConf(UEXXSIRECONF CONFTEMP)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SIRE_CONF");
                if (!oUserTable.GetByKey("0"))
                {
                    oUserTable.Code = CONFTEMP.Code;
                    oUserTable.Name = CONFTEMP.Code;
                    oUserTable.UserFields.Fields.Item("U_EXX_CONS").Value = CONFTEMP.UEXXCONS;
                    oUserTable.UserFields.Fields.Item("U_EXX_URSL").Value = CONFTEMP.UEXXURSL;
                    oUserTable.UserFields.Fields.Item("U_EXX_VSAP").Value = CONFTEMP.UEXXVSAP;
                    oUserTable.UserFields.Fields.Item("U_EXX_VSIR").Value = CONFTEMP.UEXXVSIR;
                    //oUserTable.UserFields.Fields.Item("U_EXX_APIS").Value = CONFTEMP.UEXXAPIS;
                    //oUserTable.UserFields.Fields.Item("U_EXX_USER").Value = CONFTEMP.UEXXUSER;
                    //oUserTable.UserFields.Fields.Item("U_EXX_PASS").Value = CONFTEMP.UEXXPASS;
                    //oUserTable.UserFields.Fields.Item("U_EXX_CLID").Value = CONFTEMP.UEXXCLID;
                    //oUserTable.UserFields.Fields.Item("U_EXX_CLSE").Value = CONFTEMP.UEXXCLSE;

                    if (oUserTable.Add() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }

        public static void ActualizarSireConf(UEXXSIRECONF CONFTEMP)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SIRE_CONF");
                if (oUserTable.GetByKey(CONFTEMP.Code))
                {
                    oUserTable.UserFields.Fields.Item("U_EXX_CONS").Value = CONFTEMP.UEXXCONS;
                    oUserTable.UserFields.Fields.Item("U_EXX_URSL").Value = CONFTEMP.UEXXURSL;
                    oUserTable.UserFields.Fields.Item("U_EXX_VSAP").Value = CONFTEMP.UEXXVSAP;
                    oUserTable.UserFields.Fields.Item("U_EXX_VSIR").Value = CONFTEMP.UEXXVSIR;
                    //oUserTable.UserFields.Fields.Item("U_EXX_APIS").Value = CONFTEMP.UEXXAPIS;
                    //oUserTable.UserFields.Fields.Item("U_EXX_USER").Value = CONFTEMP.UEXXUSER;
                    //oUserTable.UserFields.Fields.Item("U_EXX_PASS").Value = CONFTEMP.UEXXPASS;
                    //oUserTable.UserFields.Fields.Item("U_EXX_CLID").Value = CONFTEMP.UEXXCLID;
                    //oUserTable.UserFields.Fields.Item("U_EXX_CLSE").Value = CONFTEMP.UEXXCLSE;

                    if (oUserTable.Update() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }

        public static void CrearRegistroAPI(UEXXSIREAPIS CONFTEMP)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SIRE_APIS");
                if (!oUserTable.GetByKey("0"))
                {
                    oUserTable.Code = CONFTEMP.Code;
                    oUserTable.Name = CONFTEMP.Code;
                    oUserTable.UserFields.Fields.Item("U_EXX_RUC").Value = CONFTEMP.UEXXRUC;
                    oUserTable.UserFields.Fields.Item("U_EXX_APIS").Value = CONFTEMP.UEXXAPIS;
                    oUserTable.UserFields.Fields.Item("U_EXX_USER").Value = CONFTEMP.UEXXUSER;
                    oUserTable.UserFields.Fields.Item("U_EXX_PASS").Value = CONFTEMP.UEXXPASS;
                    oUserTable.UserFields.Fields.Item("U_EXX_CLID").Value = CONFTEMP.UEXXCLID;
                    oUserTable.UserFields.Fields.Item("U_EXX_CLSE").Value = CONFTEMP.UEXXCLSE;

                    if (oUserTable.Add() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }

        public static void ActualizarRegistroAPI(UEXXSIREAPIS CONFTEMP)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SIRE_APIS");
                if (oUserTable.GetByKey(CONFTEMP.Code))
                {
                    oUserTable.UserFields.Fields.Item("U_EXX_RUC").Value = CONFTEMP.UEXXRUC;
                    oUserTable.UserFields.Fields.Item("U_EXX_APIS").Value = CONFTEMP.UEXXAPIS;
                    oUserTable.UserFields.Fields.Item("U_EXX_USER").Value = CONFTEMP.UEXXUSER;
                    oUserTable.UserFields.Fields.Item("U_EXX_PASS").Value = CONFTEMP.UEXXPASS;
                    oUserTable.UserFields.Fields.Item("U_EXX_CLID").Value = CONFTEMP.UEXXCLID;
                    oUserTable.UserFields.Fields.Item("U_EXX_CLSE").Value = CONFTEMP.UEXXCLSE;

                    if (oUserTable.Update() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }

        public static List<UEXXSETUP> ConsultarSetup()
        {
            try
            {
                if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_ListarSetup;
                else Globals.Query = Properties.Resources.SQL_ListarSetup;

                List<UEXXSETUP> list = new List<UEXXSETUP>();
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                while (!Globals.oRec.EoF)
                {
                    list.Add(new UEXXSETUP
                    {
                        Code = Globals.oRec.Fields.Item("Code").Value?.ToString(),
                        Name = Globals.oRec.Fields.Item("Name").Value?.ToString(),
                        U_EXX_ADDN = Globals.oRec.Fields.Item("U_EXX_ADDN").Value?.ToString(),
                        U_EXX_VERS = Globals.oRec.Fields.Item("U_EXX_VERS").Value?.ToString(),
                        U_EXX_RUTA = Globals.oRec.Fields.Item("U_EXX_RUTA").Value?.ToString(),
                    });
                    Globals.oRec.MoveNext();
                }
                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void RegistrarSetup(UEXXSETUP setup)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SETUP");
                if (!oUserTable.GetByKey("0"))
                {
                    oUserTable.Code = setup.Code;
                    oUserTable.Name = setup.Name;
                    oUserTable.UserFields.Fields.Item("U_EXX_ADDN").Value = setup.U_EXX_ADDN ?? "";
                    oUserTable.UserFields.Fields.Item("U_EXX_VERS").Value = setup.U_EXX_VERS ?? "";
                    oUserTable.UserFields.Fields.Item("U_EXX_RUTA").Value = setup.U_EXX_RUTA ?? "";

                    if (oUserTable.Add() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }

        public static void ActualizarSETUP(UEXXSETUP setup)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            try
            {
                oUserTable = Globals.oCompany.UserTables.Item("EXX_SETUP");
                if (oUserTable.GetByKey(setup.Code))
                {
                    oUserTable.UserFields.Fields.Item("U_EXX_ADDN").Value = setup.U_EXX_ADDN;
                    oUserTable.UserFields.Fields.Item("U_EXX_VERS").Value = setup.U_EXX_VERS;
                    oUserTable.UserFields.Fields.Item("U_EXX_RUTA").Value = setup.U_EXX_RUTA;

                    if (oUserTable.Update() != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception(Globals.sErrMsg);
                    }
                    Globals.Release(oUserTable);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(oUserTable);
            }
        }
    }
}