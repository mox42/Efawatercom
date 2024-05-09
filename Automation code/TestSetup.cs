using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.Spreadsheet;
using OpenQA.Selenium.Support.UI;

namespace MoFinalProj
{
    internal class TestSetup
    {

        public static IWebDriver driver = new ChromeDriver();
        public static string url = "http://localhost:4200";
        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        public static void OpenDriver()
        {
            driver.Manage().Window.Maximize();

        }

        public static void AdminPage()
        {
            driver.Navigate().GoToUrl("http://localhost:4200/admin/home");

        }

        public static void ReportPage()
        {
            driver.Navigate().GoToUrl("http://localhost:4200/admin/report");

        }
        public static void NavigateToURL()
        {
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(url);
        }

        public static void CloseDriver()
        {
            driver.Quit();
        }

        public static void HighlightElement(IWebElement element)
        {
            // before click
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;


            executor.ExecuteScript("arguments[0].setAttribute('style', 'border: 2px solid lightskyblue; border-radius: 10px; background-color: lightskyblue;!important')", element);
            Thread.Sleep(1000);
            executor.ExecuteScript("arguments[0].setAttribute('style' , 'background: none !important')", element);
        }

        public static void CheckElement(IWebElement element) 
        {
            // to verfiy expected and actual result (no click)
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;


            executor.ExecuteScript("arguments[0].setAttribute('style', 'border: 2px solid lightpink; border-radius: 10px; background-color: lightpink;!important')", element);
            Thread.Sleep(1000);
            executor.ExecuteScript("arguments[0].setAttribute('style' , 'background: none !important')", element);

        }

        public static void ScrollElement(IWebElement element)
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].scrollIntoView(true);", element);
        }

        public static string TakeScreenShot()
        {
            ITakesScreenshot TakeScreenShot = (ITakesScreenshot)driver;
            Screenshot screenshot = TakeScreenShot.GetScreenshot();
            string path = "D:\\Files\\QAPractical\\MoFinalProj\\Screen\\";
            string imageName = Guid.NewGuid().ToString() + "_image.png";
            string fullPath = Path.Combine(path + $"\\{imageName}");
            screenshot.SaveAsFile(fullPath);
            return fullPath;
        }

        public static Worksheet ReadExcel(string sheetName)
        {
            Spreadsheet Excel = new Spreadsheet();
            Excel.LoadFromFile("D:\\Files\\QAPractical\\MoFinalProj\\MoData.xlsx");
            Worksheet sheet = Excel.Workbook.Worksheets.ByName(sheetName);
            return sheet;

        }


    }
}
