using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WopiHost.Webv2.Utility
{
    public interface IHttpClientHelperBase
    {
        Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, Data content = null) where T : class where Data : class;
        Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, HttpContext httpContext, Data content = null) where T : class where Data : class;
        Task<HttpResponseMessage> SendRequest(string endPoint, HttpMethod method, Dictionary<string, string> headers = null, dynamic content = null);
    }
    public class HttpClientHelperBase : IHttpClientHelperBase
    {
        private readonly IHttpClientFactory _httpClientFactory;
        public HttpClientHelperBase(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
        }
        public async Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, Data content = null) where T : class where Data : class
        {
            try
            {
                using var response = await SendRequest(endPoint, method, headers, content);
                if (response.IsSuccessStatusCode)
                {
                    var reader = await response.Content.ReadAsStringAsync();
                    var res = JsonConvert.DeserializeObject<T>(!reader.Contains("empty") ? reader : "");
                    return res;
                }
                else
                {
                    return await HandleHttpClientErrorResponse<T>(response);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, HttpContext httpContext, Data content = null) where T : class where Data : class
        {
            try
            {
                Microsoft.Extensions.Primitives.StringValues token = string.Empty;
                httpContext?.Request?.Headers?.TryGetValue("Authorization", out token);
                if (!string.IsNullOrEmpty(token.ToString()))
                {
                    if (headers == null)
                    {
                        headers = new Dictionary<string, string>();
                    }
                    headers.Add("Authorization", token);
                }
                using var response = await SendRequest(endPoint, method, headers, content);
                if (response.IsSuccessStatusCode)
                {
                    var reader = await response.Content.ReadAsStringAsync();
                    var res = JsonConvert.DeserializeObject<T>(!reader.Contains("empty") ? reader : "");
                    return res;
                }
                else
                {
                    return await HandleHttpClientErrorResponse<T>(response);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public async Task<HttpResponseMessage> SendRequest(string endPoint, HttpMethod method, Dictionary<string, string> headers = null, dynamic content = null)
        {
            const string Application_Json = "application/json";
            HttpResponseMessage response = null;
            using (HttpClient client = _httpClientFactory.CreateClient())
            {
                method ??= HttpMethod.Get;
                using var request = new HttpRequestMessage(method, endPoint);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(Application_Json));
                if (headers != null)
                {
                    foreach (var header in headers)
                    {
                        request.Headers.Add(header.Key, header.Value);
                    }
                }
                if (content != null)
                {
                    string c;
                    if (content is string)
                    {
                        c = content;
                    }
                    else
                    {
                        c = JsonConvert.SerializeObject(content);
                    }
                    request.Content = new StringContent(c, Encoding.UTF8, Application_Json);
                }
                response = await client.SendAsync(request);
            }
            return await Task.FromResult(response);
        }
        private async Task<T> HandleHttpClientErrorResponse<T>(HttpResponseMessage response) where T : class
        {
            System.Type tType = typeof(T);
            object newTInstance = Activator.CreateInstance(tType);//create an instance of T class
            PropertyInfo propertyTMessage = tType?.GetProperty("Message");
            PropertyInfo propertyTStatusCode = tType?.GetProperty("StatusCode");
            propertyTStatusCode?.SetValue(newTInstance, response.StatusCode);
            propertyTMessage?.SetValue(newTInstance, response?.ReasonPhrase);
            return await System.Threading.Tasks.Task.FromResult(newTInstance as T);
        }
    }
}
