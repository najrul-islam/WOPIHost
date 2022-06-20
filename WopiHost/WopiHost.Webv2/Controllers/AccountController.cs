using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using WopiHost.Data.Models;
using WopiHost.Data.ViewModel;
using System.Net.Http;
using WopiHost.Webv2.Utility;
using Newtonsoft.Json;
using System.Reflection;
using WopiHost.Abstractions;
using Microsoft.Extensions.Options;

namespace WopiHost.Webv2.Controllers
{
    [Authorize]
    public class AccountController : Controller
    {
        private readonly IHttpClientHelperBase _httpClientHelperBase;
        private readonly IOptionsSnapshot<WopiHostOptions> _optionsSnapshot;
        public AccountController(IHttpClientHelperBase httpClientHelperBase,
            IOptionsSnapshot<WopiHostOptions> optionsSnapshot)
        {
            _httpClientHelperBase = httpClientHelperBase;
            _optionsSnapshot = optionsSnapshot;
        }
        [HttpGet]
        [AllowAnonymous]
        // GET: AccountController
        public ActionResult Login()
        {
            /*if (!string.IsNullOrEmpty(tokenResult.Access_token))
            {
                HttpContext.Session.SetString("Token", tokenResult.Access_token);
                return RedirectToAction("Index", "Home");
            }*/
            return View();
        }
        [HttpPost]
        [AllowAnonymous]
        // POST: AccountController
        public async Task<ActionResult> Login(TokenInfo tokenInfo)
        {
            string endPoint = $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/Authorization/Token";
            var tokenResult = await _httpClientHelperBase.RequestDataAsync<TokenResult, TokenInfo>(endPoint, HttpMethod.Get, null, tokenInfo);
            if (!string.IsNullOrEmpty(tokenResult.Access_token))
            {
                HttpContext.Session.SetString("Token", tokenResult.Access_token);
                return RedirectToAction("Index", "Home");
            }
            ViewData["ErrorMessage"] = "Invalid Email or Password.";
            return View();
        }
        [HttpGet]
        //[AllowAnonymous]
        // GET: AccountController
        public ActionResult Logout()
        {
            HttpContext.Session.Remove("Token");
            return RedirectToAction("Login", "Account");
        }
        [HttpGet]
        //[AllowAnonymous]
        // GET: AccountController
        public ActionResult SignUp()
        {
            /*if (!string.IsNullOrEmpty(tokenResult.Access_token))
            {
                HttpContext.Session.SetString("Token", tokenResult.Access_token);
                return RedirectToAction("Index", "Home");
            }*/
            return View();
        }
        [HttpPost]
        [AllowAnonymous]
        // POST: AccountController
        public async Task<ActionResult> SignUp(User user)
        {
            string endPoint = $"{_optionsSnapshot.Value.WopiHostUrl}/wopi/Authorization/AddUser";
            var saveUser = await _httpClientHelperBase.RequestDataAsync<User, User>(endPoint, HttpMethod.Post, null, user);
            return RedirectToAction("Login", "Account");
        }
        /*#region Private
        private async Task<T> RequestDataAsync<T, Data>(string endPoint, HttpMethod method, Dictionary<string, string> headers, Data content = null) where T : class where Data : class
        {
            try
            {
                using var response = await _httpClientHelperBase.SendRequest(endPoint, method, headers, content);
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
        #endregion*/
    }
}
