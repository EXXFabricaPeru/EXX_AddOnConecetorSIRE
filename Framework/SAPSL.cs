using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using AddOnConectorSIRE.Entities;

namespace AddOnConectorSIRE.Framework
{
    public class SAPSL
    {
        public static string Login(UEXXSIRECONF CONF)
        {
            try
            {
                string sConnectionContextAux = null;
                try
                {
                    sConnectionContextAux = Globals.SBO_Application.Company.GetServiceLayerConnectionContext(CONF.UEXXURSL);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                if (sConnectionContextAux == null)
                    throw new Exception("No se logró establecer conexión con Service Layer");

                string sConnectionContext = sConnectionContextAux;
                return sConnectionContext.Split(';')[0].Replace("B1SESSION=", "");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void Logout()
        {
            try
            {
                if (!string.IsNullOrEmpty(Globals.B1SESSION))
                {
                    ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                    var client = new RestClient(Globals.CONF.UEXXURSL);
                    var request = new RestRequest("Logout", Method.POST);
                    request.AddHeader("content-type", "application/json");
                    request.AddCookie("B1SESSION", Globals.B1SESSION);
                    //request.AddCookie("ROUTEID", ".node0");
                    request.AddParameter("application/json", null, ParameterType.RequestBody);
                    var response = client.Execute(request);
                    if (!response.IsSuccessful) throw new InvalidOperationException(response.StatusDescription);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static List<T> Find<T>(params Tuple<string, string, string>[] prms) where T : class, new()
        {
        band:
            try
            {
                var @object = new T();
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object), Method.GET);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                string filterQuery = Globals.BuildFilter(prms);
                if (!string.IsNullOrEmpty(filterQuery))
                    request.AddParameter("$filter", filterQuery, ParameterType.QueryString);
                var rsp = client.Execute(request);

                if (rsp.StatusCode == HttpStatusCode.NotFound) return null;
                if (rsp.StatusCode != HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new ApplicationException(rsp.Content);

                JObject responseContent = JObject.Parse(rsp.Content);
                JToken values = responseContent["value"];

                if (values == null)
                    throw new Exception($"Bad Service Layer response:{rsp.Content}");

                return values.ToObject<List<T>>();
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                if (ex.Message.Contains("EntitySet"))
                    goto band;

                var findResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(findResponse.Error.Message.value);
            }
        }

        public static T FindById<T>(T @object) where T : class, new()
        {
        band:
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object), Method.GET);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                var rsp = client.Execute(request);

                if (rsp.StatusCode == HttpStatusCode.NotFound) return null;
                if (rsp.StatusCode != HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new ApplicationException(rsp.Content);

                JObject responseContent = JObject.Parse(rsp.Content);
                return responseContent.ToObject<T>();
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var findResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(findResponse.Error.Message.value);
            }
        }

        public static List<T> Find<T>(dynamic @object) where T : class, new()
        {
        band:
            try
            {
                var objetoT = new T();
                Globals.OnlyURI = true;
                string json = JsonConvert.SerializeObject(@object);
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(objetoT), Method.POST);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                var rsp = client.Execute(request);

                if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                    throw new ApplicationException(rsp.Content);

                JObject responseContent = JObject.Parse(rsp.Content);
                JToken values = responseContent["value"];

                if (values == null)
                    throw new Exception($"Bad Service Layer response:{rsp.Content}");

                return values.ToObject<List<T>>();
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var createResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(createResponse.Error.Message.value);
            }
        }

        public static string Add<T>(T @object)
        {
        band:
            try
            {
                Globals.OnlyURI = true;
                string json = JsonConvert.SerializeObject(@object, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object), Method.POST);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                var rsp = client.Execute(request);

                if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                    throw new ApplicationException(rsp.Content);

                return rsp.Content;
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var createResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(createResponse.Error.Message.value);
            }
        }

        public static string Update<T>(T @object)
        {
        band:
            try
            {
                string json = JsonConvert.SerializeObject(@object, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object), Method.PATCH);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                var rsp = client.Execute(request);

                if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                    throw new ApplicationException(rsp.Content);
                return rsp.Content;
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var updateResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(updateResponse.Error.Message.value);
            }
        }

        public static string Delete<T>(T @object)
        {
        band:
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object), Method.DELETE);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                var rsp = client.Execute(request);

                if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                    throw new ApplicationException(rsp.Content);
                return rsp.Content;
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var deleteResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(deleteResponse.Error.Message.value);
            }
        }

        public static string Cancel<T>(T @object)
        {
        band:
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                var client = new RestClient(Globals.CONF.UEXXURSL);
                var request = new RestRequest(Globals.BuildURI(@object) + "/Cancel", Method.POST);
                request.AddHeader("content-type", "application/json");
                request.AddCookie("B1SESSION", Globals.B1SESSION);
                var rsp = client.Execute(request);

                if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                    throw new ApplicationException(rsp.Content);
                return rsp.Content;
            }
            catch (Exception ex)
            {
                if (ex.Message == "No se logró establecer conexión con Service Layer")
                    throw ex;

                if (ex.Message.Contains("Invalid session"))
                {
                    Globals.B1SESSION = Login(Globals.CONF);
                    goto band;
                }

                var deleteResponse = JsonConvert.DeserializeObject<ErrorResponse>(ex.Message);
                throw new Exception(deleteResponse.Error.Message.value);
            }
        }
    }
}