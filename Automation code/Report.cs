using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoFinalProj
{
    internal class Report
    {


        // CONFIGURATION
        // -------------------------------------------------------------------------------------------
        public Users user { get; set; }
        public Report(Users user)
        {
            this.user = user;
        }



        // Strings
        // -------------------------------------------------------------------------------------------

        public static string ReportNameString()
        {
            string case1 = "The test case aims to verify the functionality of searching By Category Name on the report page";

            return case1;
        }

        public static string ReportDateString()
        {
            string case1 = "The test case aims to verify the functionality of searching & filtering by date, " +
                "and confirming that the payment history dates fall within the specified date interval. ";

            return case1;
        }

        public static string ReportCountString()
        {
            string case1 = "The test case aims to verify that the number of payments on Report page matches the count of payments history in payment history page";

            return case1;
        }




        // REPORT ELEMENTS
        // -------------------------------------------------------------------------------------------
        public void StartReport()
        {
            Thread.Sleep(500);
            IWebElement ReportSection = TestSetup.driver.FindElement(By.XPath("//a[@class='nav-link' and .//i[@class='bi bi-calculator']]"));
            TestSetup.ScrollElement(ReportSection);
            TestSetup.HighlightElement(ReportSection);
            ReportSection.Click();

        }

        public void ReportDropdownList()
        {

            Thread.Sleep(2000);
            IWebElement ReportDropdownList = TestSetup.driver.FindElement(By.XPath("//select[@formcontrolname='Billercategoryname']"));
            TestSetup.HighlightElement(ReportDropdownList);

            ReportDropdownList.SendKeys(Keys.Right);

            // Get all the options within the dropdown
            List<IWebElement> options = new List<IWebElement>(ReportDropdownList.FindElements(By.TagName("option")));

            // Generate a random index within the range of available options
            Random random = new Random();
            int randomIndex = random.Next(0, options.Count);

            // Select the option at the randomly generated index
            options[randomIndex].Click();
            Thread.Sleep(1000);

        }

        public void ReportDate(string date1 , string date2)
        {

            IWebElement ReportDateFrom = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='DateFrom']"));
            TestSetup.HighlightElement(ReportDateFrom);
            ReportDateFrom.SendKeys(date1);
            ReportDateFrom.SendKeys(Keys.ArrowUp);
            ReportDateFrom.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);

            IWebElement ReportDateTo = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='DateTo']"));
            TestSetup.HighlightElement(ReportDateTo);
            ReportDateTo.SendKeys(date2);
            ReportDateTo.SendKeys(Keys.ArrowUp);
            ReportDateTo.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);
        }

        public void ClearDateInputField()
        {

            IWebElement ClearDateFrom = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='DateFrom']"));
            ClearDateFrom.Clear();
            IWebElement ClearDateTo = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='DateTo']"));
            ClearDateTo.Clear();

        }

        public void ClearDropDownList()
        {
            IWebElement ReportDropdownList = TestSetup.driver.FindElement(By.XPath("//select[@formcontrolname='Billercategoryname']"));
            // Use JavaScript to reset the dropdown to its default state
            ((IJavaScriptExecutor)TestSetup.driver).ExecuteScript("arguments[0].selectedIndex = -1;", ReportDropdownList);
        }





    }
}
