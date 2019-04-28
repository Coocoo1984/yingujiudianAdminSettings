using System;
using System.Data;

namespace BasicSettingsMVC.Models
{
    public class LoginModel
    {
        string userName;
        string password;

        public string UserName { get => userName; set => userName = value; }
        public string Password { get => password; set => password = value; }

        public bool Login(string username, string password)
        {
            bool result = false;
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
            }
            else
            {
                if (username == "admin" && password == "ygjd@1234")
                {
                    result = true;
                }
            }
            return result;
        }


    }
}
