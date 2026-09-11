using AddOnConectorSIRE.Entities;
using AddOnConectorSIRE.Utilities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AddOnConectorSIRE.Framework
{
    public class SIREAPI
    {
        public static string GenerarToken(UEXXSIREAPIS empresa)
        {
            try
            {
                try
                {
                    ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

                    var client = new RestClient("https://api-seguridad.sunat.gob.pe");
                    var request = new RestRequest($"/v1/clientessol/{empresa.UEXXCLID}/oauth2/token/", Method.POST);

                    request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

                    request.AddParameter("grant_type", "password");
                    request.AddParameter("scope", empresa.UEXXAPIS);
                    request.AddParameter("client_id", empresa.UEXXCLID);
                    request.AddParameter("client_secret", empresa.UEXXCLSE);
                    request.AddParameter("username", empresa.UEXXRUC + empresa.UEXXUSER);
                    request.AddParameter("password", ExxisEncryptor.Decrypt(empresa.UEXXPASS));

                    var rsp = client.Execute(request);

                    if (rsp.StatusCode != HttpStatusCode.OK)
                        throw new ApplicationException("Error obteniendo token SUNAT: " + rsp.Content);

                    var tokenResponse = JsonConvert.DeserializeObject<TOKEN_SUNAT>(rsp.Content);
                    return tokenResponse.access_token;
                }
                catch (Exception ex)
                {
                    throw (ex);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static T ConsultarDocumento<T>(string oEmpresa, string uri)
        {
        band:
            try
            {
                var empresa = Globals.CONF.APIS.Where(x => x.Code == oEmpresa).ToList();
                if (empresa.Count == 0)
                    throw new Exception("No existe configuración de empresa.");

                if (string.IsNullOrEmpty(empresa[0].TOKEN)) empresa[0].TOKEN = GenerarToken(empresa[0]);

                //string json = JsonConvert.SerializeObject(osunat);
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

                var client = new RestClient(empresa[0].UEXXAPIS);
                var request = new RestRequest(uri, Method.GET);

                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Authorization", $"Bearer {empresa[0].TOKEN}");

                //request.AddJsonBody(json);

                var rsp = client.Execute(request);

                if (rsp.StatusCode == HttpStatusCode.NotFound)
                    throw new Exception("No se logró obtener respuesta del servicio de SUNAT.");

                if (rsp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    empresa[0].TOKEN = GenerarToken(empresa[0]);
                    goto band;
                }

                if (rsp.StatusCode != HttpStatusCode.OK)
                    throw new ApplicationException(rsp.Content);

                if (typeof(T) == typeof(string))
                    return (T)(object)rsp.Content;

                if (EsJsonValido(rsp.Content))
                    return JsonConvert.DeserializeObject<T>(rsp.Content);

                if (typeof(T) == typeof(byte[]))
                    return (T)(object)rsp.RawBytes;

                return (T)(object)rsp.Content;
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        public static T DescargarResumen<T>(string oEmpresa, string uri)
        {
        band:
            try
            {
                var empresa = Globals.CONF.APIS.Where(x => x.Code == oEmpresa).ToList();
                if (empresa.Count == 0)
                    throw new Exception("No existe configuración de empresa.");

                if (string.IsNullOrEmpty(empresa[0].TOKEN)) empresa[0].TOKEN = GenerarToken(empresa[0]);

                //string json = JsonConvert.SerializeObject(osunat);
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

                var client = new RestClient(empresa[0].UEXXAPIS);
                var request = new RestRequest(uri, Method.GET);

                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Authorization", $"Bearer {empresa[0].TOKEN}");

                //request.AddJsonBody(json);

                var rsp = client.Execute(request);

                if (rsp.StatusCode == HttpStatusCode.NotFound)
                    throw new Exception("No se logró obtener respuesta del servicio de SUNAT.");

                if (rsp.StatusCode == HttpStatusCode.Unauthorized)
                {
                    empresa[0].TOKEN = GenerarToken(empresa[0]);
                    goto band;
                }

                if (rsp.StatusCode != HttpStatusCode.OK)
                    throw new ApplicationException(rsp.Content);

                if (typeof(T) == typeof(string))
                    return (T)(object)rsp.Content;

                if (EsJsonValido(rsp.Content))
                    return JsonConvert.DeserializeObject<T>(rsp.Content);

                if (typeof(T) == typeof(byte[]))
                    return (T)(object)rsp.RawBytes;

                return (T)(object)rsp.Content;
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        public static string EnviarDocumento(string oEmpresa, string rutaZip, string perTributario, string libro, string proceso, string url)
        {
        band:
            try
            {
                var empresa = Globals.CONF.APIS
                    .Where(x => x.Code == oEmpresa)
                    .ToList();

                if (empresa.Count == 0) throw new Exception("No existe configuración de empresa.");
                if (string.IsNullOrEmpty(empresa[0].TOKEN)) empresa[0].TOKEN = GenerarToken(empresa[0]);

                FileInfo fileInfo = new FileInfo(rutaZip);
                if (!fileInfo.Exists) throw new Exception("No existe el archivo ZIP.");
                string fileName = fileInfo.Name;

                byte[] fileBytes = File.ReadAllBytes(rutaZip);

                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                string metadata = "filename " + Convert.ToBase64String(Encoding.UTF8.GetBytes(fileName)) +
                                  ",filetype " + Convert.ToBase64String(Encoding.UTF8.GetBytes("application/zip")) +
                                  ",numRuc " + Convert.ToBase64String(Encoding.UTF8.GetBytes(empresa[0].UEXXRUC)) +
                                  ",perTributario " + Convert.ToBase64String(Encoding.UTF8.GetBytes(perTributario)) +
                                  ",codOrigenEnvio " + Convert.ToBase64String(Encoding.UTF8.GetBytes("1")) +
                                  ",codProceso " + Convert.ToBase64String(Encoding.UTF8.GetBytes(proceso)) +
                                  ",codTipoCorrelativo " + Convert.ToBase64String(Encoding.UTF8.GetBytes("01")) +
                                  ",nomArchivoImportacion " + Convert.ToBase64String(Encoding.UTF8.GetBytes(fileName)) +
                                  ",codLibro " + Convert.ToBase64String(Encoding.UTF8.GetBytes(libro));

                var client = new RestClient(empresa[0].UEXXAPIS);
                client.Timeout = 1000 * 60 * 60 * 2;
                var requestCreate = new RestRequest(url, Method.POST);

                requestCreate.AddHeader("Authorization", $"Bearer {empresa[0].TOKEN}");
                requestCreate.AddHeader("Tus-Resumable", "1.0.0");
                requestCreate.AddHeader("Upload-Length", fileBytes.Length.ToString());
                requestCreate.AddHeader("Upload-Metadata", metadata);
                requestCreate.AddHeader("Content-Type", "application/offset+octet-stream");

                var responseCreate = client.Execute(requestCreate);

                if (responseCreate.StatusCode == HttpStatusCode.Unauthorized)
                {
                    empresa[0].TOKEN = GenerarToken(empresa[0]);
                    goto band;
                }

                if (!responseCreate.IsSuccessful)
                    throw new Exception("Error creando upload TUS: " + responseCreate.StatusCode + " - " + responseCreate.Content);

                string uploadUrl =
                    responseCreate.Headers
                    .Where(x => x.Name == "Location")
                    .Select(x => x.Value.ToString())
                    .FirstOrDefault();

                if (string.IsNullOrEmpty(uploadUrl))
                    throw new Exception("SUNAT no retornó URL TUS.");


                var uploadClient = new RestClient(empresa[0].UEXXAPIS);
                uploadClient.Timeout = 1000 * 60 * 60 * 2;
                var requestPatch = new RestRequest(uploadUrl, Method.PATCH);

                requestPatch.AddHeader("Authorization", $"Bearer {empresa[0].TOKEN}");
                requestPatch.AddHeader("Tus-Resumable", "1.0.0");
                requestPatch.AddHeader("Upload-Offset", "0");
                requestPatch.AddHeader("Content-Type", "application/offset+octet-stream");
                requestPatch.AddParameter("application/offset+octet-stream", fileBytes, ParameterType.RequestBody);

                var responsePatch = uploadClient.Execute(requestPatch);

                if (responsePatch.StatusCode == HttpStatusCode.Unauthorized)
                {
                    empresa[0].TOKEN = GenerarToken(empresa[0]);
                    goto band;
                }

                if (!responsePatch.IsSuccessful)
                    throw new Exception("Error subiendo archivo: " + responsePatch.StatusCode + " - " + responsePatch.Content);

                return responsePatch.Content;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static string ToBase64(string value)
        {
            return Convert.ToBase64String(
                Encoding.UTF8.GetBytes(value));
        }

        private static bool EsJsonValido(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return false;

            str = str.Trim();

            if ((str.StartsWith("{") && str.EndsWith("}")) ||
                (str.StartsWith("[") && str.EndsWith("]")))
            {
                try
                {
                    JToken.Parse(str);
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }
    }
}
