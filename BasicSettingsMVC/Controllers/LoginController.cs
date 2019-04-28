using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using BasicSettingsMVC.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc;

namespace BasicSettingsMVC.Controllers
{
    public class LoginController : Controller
    {
        // GET: /<controller>/
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Login(string userName, string password)
        {
            if (new LoginModel().Login(userName, password))
            {
                var claims = new List<Claim>()
                {
                    new Claim(ClaimTypes.Name,userName)
                    //,new Claim(ClaimTypes.Email,"emailaccount@microsoft.com")  
                };

                //init the identity instances 
                var userPrincipal = new ClaimsPrincipal(new ClaimsIdentity(claims, "SuperSecureLogin"));

                //signin 
                HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, userPrincipal);

                return RedirectToAction("Index", "Home");
            }
            else
            {
                ViewBag.ErrMsg = "账号或密码无效";

                return View();
            }
        }

        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

            return RedirectToAction("Index", "Login");
        }

    }
}
