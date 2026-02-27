using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Framework;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace AddOnConectorSIRE
{
    public class Setup
    {
        public static void CargarConfiguracion()
        {
            try
            {
                if (Globals.IsHana()) Globals.Query = Properties.Resources.HANA_MultiBranch;
                else Globals.Query = Properties.Resources.SQL_MultiBranch;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();
                Globals.isMultiBranch = Convert.ToInt32(Globals.oRec.Fields.Item(0).Value) > 0 ? true : false;
                Globals.Release(Globals.oRec);

                if (Globals.IsHana())
                    Globals.Query = Properties.Resources.HANA_ValidarTabla;
                else
                    Globals.Query = Properties.Resources.SQL_ValidarTabla;

                Globals.Query = string.Format(Globals.Query, Globals.oCompany.CompanyDB, "@EXX_SIRE_CONF");
                Globals.RunQuery(Globals.Query);

                if (Globals.oRec.RecordCount == 0)
                {
                    Globals.Release(Globals.oRec);
                    Globals.InformationMessage($"No existe tabla 'EXX_SIRE_CONF', se creará la tabla para cargar la configuración inicial del {Globals.AddOnName}");
                    CreateTablesConf();
                    CreateFieldsConf();
                    for (int i = Globals.SBO_Application.Forms.Count - 1; i >= 0; i--)
                    {
                        SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(i);

                        if (oForm.Modal)
                            oForm.Items.Item("1").Click();
                    }
                    Globals.InformationMessage($"Creación de tabla de configuración terminó correctamente, por favor llene la configuración inicial del {Globals.AddOnName} en: Modulos > EXX-SIRE SUNAT > EXX-CONFIGURACION SIRE");
                }
                else
                {
                    Globals.Release(Globals.oRec);

                    if (Globals.IsHana())
                        Globals.Query = Properties.Resources.HANA_ObtenerConfiguracion;
                    else
                        Globals.Query = Properties.Resources.SQL_ObtenerConfiguracion;
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.RecordCount > 0)
                    {
                        Globals.oRec.MoveFirst();
                        UEXXSIRECONF CONFTEMP = new Entities.UEXXSIRECONF();
                        CONFTEMP.Code = Globals.oRec.Fields.Item("Code").Value.ToString();
                        CONFTEMP.UEXXCONS = Globals.oRec.Fields.Item("U_EXX_CONS").Value.ToString();
                        CONFTEMP.UEXXURSL = Globals.oRec.Fields.Item("U_EXX_URSL").Value.ToString();
                        CONFTEMP.UEXXVSAP = Globals.oRec.Fields.Item("U_EXX_VSAP").Value.ToString();
                        CONFTEMP.UEXXVSIR = Globals.oRec.Fields.Item("U_EXX_VSIR").Value.ToString();
                        Globals.Release(Globals.oRec);
                        CONFTEMP.APIS = new List<UEXXSIREAPIS>();

                        if (Globals.IsHana())
                        {
                            if (Globals.isMultiBranch) Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxSucursal;
                            else Globals.Query = Properties.Resources.HANA_ObtieneConfAPIxOADM;
                        }
                        else
                        {
                            if (Globals.isMultiBranch) Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxSucursal;
                            else Globals.Query = Properties.Resources.SQL_ObtieneConfAPIxOADM;
                        }
                        Globals.RunQuery(Globals.Query);
                        Globals.oRec.MoveFirst();
                        while (!Globals.oRec.EoF)
                        {
                            Globals.isMultiBranch = Convert.ToInt32(Globals.oRec.Fields.Item(0).Value) > 0 ? true : false;
                            CONFTEMP.APIS.Add(new UEXXSIREAPIS
                            {
                                Code = Globals.oRec.Fields.Item("BPLId").Value.ToString(),
                                UEXXRUC = Globals.oRec.Fields.Item("GlblLocNum").Value.ToString(),
                                UEXXAPIS = Globals.oRec.Fields.Item("U_EXX_APIS").Value.ToString(),
                                UEXXUSER = Globals.oRec.Fields.Item("U_EXX_USER").Value.ToString(),
                                UEXXPASS = Globals.oRec.Fields.Item("U_EXX_PASS").Value.ToString(),
                                UEXXCLID = Globals.oRec.Fields.Item("U_EXX_CLID").Value.ToString(),
                                UEXXCLSE = Globals.oRec.Fields.Item("U_EXX_CLSE").Value.ToString(),
                            });
                            Globals.oRec.MoveNext();
                        }
                        //Globals.Release(Globals.oRec);

                        //CONFTEMP.UEXXAPIS = Globals.oRec.Fields.Item("U_EXX_APIS").Value.ToString();
                        //CONFTEMP.UEXXUSER = Globals.oRec.Fields.Item("U_EXX_USER").Value.ToString();
                        //CONFTEMP.UEXXPASS = Globals.oRec.Fields.Item("U_EXX_PASS").Value.ToString();
                        //CONFTEMP.UEXXCLID = Globals.oRec.Fields.Item("U_EXX_CLID").Value.ToString();
                        //CONFTEMP.UEXXCLSE = Globals.oRec.Fields.Item("U_EXX_CLSE").Value.ToString();
                        if (Modules.Configuracion.Main.ValidaConfig(CONFTEMP))
                        {
                            Globals.CONF = new UEXXSIRECONF();
                            Globals.CONF = CONFTEMP;
                            Globals.existeConf = true;
                        }
                        return;
                    }
                    Globals.InformationMessage($"Por favor llene la configuración inicial del {Globals.AddOnName} en: Modulos > EXX-SIRE SUNAT > EXX-CONFIGURACION SIRE");
                }
            }
            catch (Exception ex)
            {
                Globals.MessageBox(ex.Message);
                Globals.existeConf = false;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ValidarVersion()
        {
            try
            {
                if (Globals.CONF.UEXXCONS == "2") Globals.B1SESSION = SAPSL.Login(Globals.CONF);
                List<UEXXSETUP> SetupList = new List<UEXXSETUP>();
                try
                {
                    if (Globals.CONF.UEXXCONS == "1") SetupList = SAPSDK.ConsultarSetup();
                    else SetupList = SAPSL.Find<UEXXSETUP>();
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                var setup = SetupList.FirstOrDefault(x => x.U_EXX_ADDN == Assembly.GetEntryAssembly().GetName().Name);
                Version actual = new Version(Assembly.GetEntryAssembly().GetName().Version.ToString());

                if (setup == null)
                {
                    Globals.InformationMessage($"Iniciando creación de metadata en {Globals.oCompany.CompanyDB}.");

                    CreateTables();
                    CreateFields();
                    CreateUDOs();
                    //Falta crear el procedimiento en automatico por SL no se puede

                    string Code = string.Empty;
                    if (SetupList.Count == 0)
                        Code = "01";
                    else
                        Code = Convert.ToInt32(SetupList.Max(x => Convert.ToInt32(x.Code)) + 1).ToString().PadLeft(2, '0');
                    setup = new UEXXSETUP
                    {
                        Code = Code,
                        Name = Code,
                        U_EXX_ADDN = Assembly.GetEntryAssembly().GetName().Name,
                        U_EXX_VERS = Assembly.GetEntryAssembly().GetName().Version.ToString()
                    };

                    Globals.OnlyURI = true;
                    if (Globals.CONF.UEXXCONS == "1") SAPSDK.RegistrarSetup(setup);
                    else SAPSL.Add<UEXXSETUP>(setup);
                    Globals.InformationMessage($"Instalación finalizada en {Globals.oCompany.CompanyDB}.");
                }
                else
                {
                    Version configurado = new Version(setup.U_EXX_VERS);
                    int comparison = actual.CompareTo(configurado);
                    if (comparison != 0)
                    {
                        if (comparison > 0)
                        {
                            Globals.InformationMessage($"Iniciando actualización de metadata en {Globals.oCompany.CompanyDB}.");

                            CreateTables();
                            CreateFields();
                            CreateUDOs();

                            setup.U_EXX_VERS = actual.ToString();
                            if (Globals.CONF.UEXXCONS == "1") SAPSDK.ActualizarSETUP(setup);
                            else SAPSL.Update<UEXXSETUP>(setup);
                            Globals.InformationMessage($"Actualización finalizada en {Globals.oCompany.CompanyDB}.");
                        }
                        else
                        {
                            throw new Exception("La versión del servicio que se está ejecutando es inferior a la versión configurada en SAP, no se habilitará los módulos del Addon.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateTablesConf()
        {
            try
            {
                var ContentList = Globals.ReadJson<UserTablesMd>("TablesConf.json");
                foreach (UserTablesMd table in ContentList)
                {
                    if (SAPSDK.CrearTabla(table))
                        Globals.InformationMessage($"TABLA ({table.TableName}): creado correctamente.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateTables()
        {
            try
            {
                var ContentList = Globals.ReadJson<UserTablesMd>("Tables.json");
                foreach (UserTablesMd table in ContentList)
                {
                    if (Globals.CONF.UEXXCONS == "1")
                    {
                        if (SAPSDK.CrearTabla(table)) Globals.InformationMessage($"TABLA ({table.TableName}): creado correctamente.");
                    }
                    else
                    {
                        var existe = SAPSL.Find<UserTablesMd>(new Tuple<string, string, string>("TableName", Globals.Operador.Igual, table.TableName));

                        if (existe.Count == 0)
                        {
                            Globals.OnlyURI = true;
                            SAPSL.Add<UserTablesMd>(table);
                            Globals.InformationMessage($"TABLA ({table.TableName}): creado correctamente.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateFieldsConf()
        {
            try
            {
                var ContentList = Globals.ReadJson<UserFieldsMd>("FieldsConf.json");
                foreach (UserFieldsMd field in ContentList)
                {
                    if (SAPSDK.CrearCampo(field))
                        Globals.InformationMessage($"CAMPO ({field.Name}): creado correctamente.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateFields()
        {
            try
            {
                var ContentList = Globals.ReadJson<UserFieldsMd>("Fields.json");
                foreach (UserFieldsMd field in ContentList)
                {
                    if (Globals.CONF.UEXXCONS == "1")
                    {
                        if (SAPSDK.CrearCampo(field)) Globals.InformationMessage($"CAMPO ({field.Name}): creado correctamente.");
                    }
                    else
                    {
                        var existe = SAPSL.Find<UserFieldsMd>(new Tuple<string, string, string>("TableName ", Globals.Operador.Igual, field.TableName)
                                                                  , new Tuple<string, string, string>("Name", Globals.Operador.Igual, field.Name));

                        if (existe.Count == 0)
                        {
                        volver:
                            try
                            {
                                Globals.OnlyURI = true;
                                SAPSL.Add<UserFieldsMd>(field);
                                Globals.InformationMessage($"CAMPO ({field.Name}): creado correctamente.");
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message == "No existe tabla (ODBC -2004)")
                                {
                                    goto volver;
                                }
                                throw ex;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateUDOs()
        {
            try
            {
                var ContentList = Globals.ReadJson<UserObjectsMd>("UDOs.json");


                foreach (UserObjectsMd udo in ContentList)
                {
                    if (Globals.CONF.UEXXCONS == "1")
                    {
                        if (SAPSDK.CrearUDO(udo)) Globals.InformationMessage($"UDO ({udo.Code}): creado correctamente.");
                    }
                    else
                    {
                        var existe = SAPSL.FindById<UserObjectsMd>(udo);
                        if (existe == null)
                        {
                            Globals.OnlyURI = true;
                            SAPSL.Add<UserObjectsMd>(udo);
                            Globals.InformationMessage($"UDO ({udo.Code}): creado correctamente.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
