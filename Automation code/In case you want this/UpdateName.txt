using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoFinalProj
{
    internal class UpdateName
    {

        // CONFIGURATION
        // -------------------------------------------------------------------------------------------
        public Users user { get; set; }
        public UpdateName(Users user)
        {
            this.user = user;
        }



        // Strings
        // -------------------------------------------------------------------------------------------
        public static string ProfileString()
        {
            string case1 = "This test case is to verify the functionality of the updating account's fullName on the Account section of the website.";
            return case1;
        }



        // UPDATE NAME ELEMENTS
        // -------------------------------------------------------------------------------------------
        public void StartProfile()
        {
            Thread.Sleep(500);
            IWebElement AccountSection = TestSetup.driver.FindElement(By.XPath("//a[@class='nav-link' and .//i[@class='bi bi-person-square']]"));
            TestSetup.ScrollElement(AccountSection);
            TestSetup.HighlightElement(AccountSection);
            AccountSection.Click();
        }

        public void UpdateFullName()
        {
            IWebElement UpdateFullName = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='FullName']"));
            TestSetup.HighlightElement(UpdateFullName);
            UpdateFullName.SendKeys(user.FullName);
            Thread.Sleep(1000);

        }

        public void EnterPassword()
        {
            IWebElement Password = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-prfile-a/div/div/div/div/div/form/div/div/div[4]/div/input"));
            TestSetup.HighlightElement(Password);
            Password.SendKeys(user.Password);
        }

        public void UpdateBtn()
        {
            IWebElement UpdateBtn = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-prfile-a/div/div/div/div/div/form/div/div/div[9]/div/div[2]/div/button"));
            TestSetup.HighlightElement(UpdateBtn);
            UpdateBtn.Click();
        }

        public void ClearInputField()
        {
            IWebElement ClearFullName = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-prfile-a/div/div/div/div/div/form/div/div/div[1]/div/input"));
            ClearFullName.Clear();
        }



    }
}
