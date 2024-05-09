using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoFinalProj
{
    internal class Login
    {


        // CONFIGURATION
        // -------------------------------------------------------------------------------------------

        public Users user { get; set; }
        public Login(Users user)
        {
            this.user = user;
        }




        // Strings
        // -------------------------------------------------------------------------------------------
        public static string  LoginString()
        {
            string case1 = "This test case is designed to verify the login functionality of the website with Admin credentials.";
            return case1;
        }



        // LOGIN ELEMENTS
        // -------------------------------------------------------------------------------------------
        public void StartLogin()
        {

            TestSetup.NavigateToURL();
            IWebElement LoginBtn = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-index/app-header/header/div[3]/div/div/nav/div[2]/div[2]/div/a[1]"));
            TestSetup.HighlightElement(LoginBtn);
            LoginBtn.Click();
        }


        public void EnterUsername()
        {
            IWebElement Username = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-login/div/div/div[1]/form/input[1]"));
            TestSetup.HighlightElement(Username);
            Username.SendKeys(user.Username);
        }

        public void EnterPassword()
        {
            IWebElement Password = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-login/div/div/div[1]/form/input[2]"));
            TestSetup.HighlightElement(Password);
            Password.SendKeys(user.Password);
        }

        public void SignInBtn()
        {
            IWebElement SignInBtn = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-login/div/div/div[1]/form/button"));
            TestSetup.HighlightElement(SignInBtn);
            SignInBtn.Click();
        }


    }
}
