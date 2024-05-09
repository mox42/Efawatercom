using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoFinalProj
{
    internal class Category
    {

        // CONFIGURATION
        // -------------------------------------------------------------------------------------------

        public Users user { get; set; }
        public Category(Users user)
        {
            this.user = user;
        }



        // Strings
        // -------------------------------------------------------------------------------------------

        public static string CategoryString()
        {
            string case1 = "This test case is Focusing on verifying the functionality of creating a category on the category section of the website.";
            return case1;
        }

        public static string CompanyString()
        {
            string case1 = "This test case is to verify the functionality of creating a Company associated with a category on the category section of the website.";
            return case1;
        }

        public static string CompanyShowButtonString()
        {
            string case1 = "This test case is to verify the functionality of creating a Company associated with a category When Show Button Is pressed on the category section of the website.";
            return case1;
        }

        public static string UpdateCategoryString()
        {
            string case1 = "This test case is to verify the functionality of Updating an existing category on the category section of the website.";
            return case1;
        }

        public static string UpdateCategoryShowButtonString()
        {
            string case1 = "This test case is Focusing on verifying the functionality of Updating an existing category When Show Button Is pressed on the category section of the website.";
            return case1;
        }



        // CATEGORY ELEMENTS - ENTER
        // -------------------------------------------------------------------------------------------

        public void CreateCategory()
        {
            Thread.Sleep(1000);
            IWebElement CreateCategory = TestSetup.driver.FindElement(By.XPath("//button[text()='Create Category Name']"));
            TestSetup.ScrollElement(CreateCategory);
            TestSetup.HighlightElement(CreateCategory);
            CreateCategory.Click();

        }

        public void EnterBillname()
        {
            IWebElement EnterBillname = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='billercategoryname']"));
            TestSetup.HighlightElement(EnterBillname);
            EnterBillname.SendKeys(user.Billname);
            Thread.Sleep(1000);
        }


        public void ChooseImage(string imagePath)
        {

            IWebElement upload = TestSetup.driver.FindElement(By.XPath("//input[@accept='image/*' and @type='file' and @formcontrolname='Pictuer' and @class='col-5 ng-untouched ng-pristine ng-invalid']"));
            TestSetup.HighlightElement(upload);

            upload.SendKeys(imagePath);
            Thread.Sleep(2000);

            Actions actions = new Actions(TestSetup.driver);
            actions.SendKeys(Keys.Right).Perform();
            Thread.Sleep(1000);

        }

        public void EnterEmail()
        {
            IWebElement EnterEmail = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='Email']"));
            TestSetup.HighlightElement(EnterEmail);
            EnterEmail.SendKeys(user.Email);
            Thread.Sleep(1000);

        }


        public void EnterLocation()
        {
            IWebElement EnterLocation = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='Location']"));
            TestSetup.HighlightElement(EnterLocation);
            EnterLocation.SendKeys(user.Location);
            TestSetup.ScrollElement(EnterLocation);
            Thread.Sleep(1000);
        }

        public void EnterDeducted()
        {
            IWebElement EnterDeducted = TestSetup.driver.FindElement(By.XPath("//input[@type='number' and @formcontrolname='deducted' and @class='col-5 ng-untouched ng-pristine ng-invalid']"));
            TestSetup.HighlightElement(EnterDeducted);
            EnterDeducted.SendKeys(user.deducted);
            Thread.Sleep(1000);
        }

        public void CreateBtn()
        {
            IWebElement CreateBtn = TestSetup.driver.FindElement(By.XPath("//button[text()='Create']"));
            TestSetup.HighlightElement(CreateBtn);
            CreateBtn.Click();
        }


        // COMPANY ELEMENTS - ENTER
        // -------------------------------------------------------------------------------------------

        public void EnterCompanyBillname()
        {
            IWebElement EnterCompanyBillname = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='Billname']"));
            TestSetup.HighlightElement(EnterCompanyBillname);
            EnterCompanyBillname.SendKeys(user.CompanyBillname);
            Thread.Sleep(1000);
        }

        public void CreateCompanyBtn()
        {
            IWebElement CreateCompanyBtn = TestSetup.driver.FindElement(By.XPath("//div[4]/button"));
            TestSetup.HighlightElement(CreateCompanyBtn);
            CreateCompanyBtn.Click();
        }


        // COMPANY ELEMENTS - UPDATE
        // -------------------------------------------------------------------------------------------

        public void UpdatBillname()
        {
            IWebElement UpdatBillname = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='billercategoryname']"));
            TestSetup.HighlightElement(UpdatBillname);
            UpdatBillname.SendKeys(user.Billname);
            Thread.Sleep(1000);
        }

        public void UpdateEmail()
        {
            IWebElement UpdateEmail = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='email']"));
            TestSetup.HighlightElement(UpdateEmail);
            UpdateEmail.SendKeys(user.Email);
            Thread.Sleep(1000);

        }

        public void UpdatLocation()
        {
            IWebElement UpdatLocation = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='location']"));
            TestSetup.HighlightElement(UpdatLocation);
            UpdatLocation.SendKeys(user.Location);
            TestSetup.ScrollElement(UpdatLocation);
            Thread.Sleep(1000);
        }

        public void UpdatDeducted()
        {
            IWebElement UpdatDeducted = TestSetup.driver.FindElement(By.XPath("//input[@type='number' and @formcontrolname='deducted']"));
            TestSetup.HighlightElement(UpdatDeducted);
            UpdatDeducted.SendKeys(user.deducted);
            Thread.Sleep(1000);
        }


        public void UpdateImage(string imagePath)
        {

            IWebElement upload = TestSetup.driver.FindElement(By.XPath("//input[@accept='image/*']"));
            TestSetup.HighlightElement(upload);

            upload.SendKeys(imagePath);
            Thread.Sleep(2000);

            Actions actions = new Actions(TestSetup.driver);
            actions.SendKeys(Keys.Right).Perform();
            Thread.Sleep(1000);
        }

        public void UpdatBtn()
        {
            IWebElement UpdatBtn = TestSetup.driver.FindElement(By.XPath("//button[text()='Update']"));
            TestSetup.HighlightElement(UpdatBtn);
            UpdatBtn.Click();
            Thread.Sleep(2000);
        }

        public void ClearUpdateInputField()
        {
            Thread.Sleep(1000);
            IWebElement UpdatBillname = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='billercategoryname']"));
            UpdatBillname.Clear();
            IWebElement UpdateEmail = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='email']"));
            UpdateEmail.Clear();
            IWebElement UpdatLocation = TestSetup.driver.FindElement(By.XPath("//input[@formcontrolname='location']"));
            UpdatLocation.Clear();
            IWebElement UpdatDeducted = TestSetup.driver.FindElement(By.XPath("//input[@type='number' and @formcontrolname='deducted']"));
            UpdatDeducted.Clear();
        }


        // VIERIFY IMAGE ELEMENT
        // -------------------------------------------------------------------------------------------
        public static void ImagePresentProcess(int index, ref string Image)
        {
            // Get WebElement reference of logo
            try
            {
                IWebElement logoElement = TestSetup.driver.FindElement(By.XPath($"/html/body/app-root/app-home/div/div/div/table/tbody[{index}]/tr/td[2]/img"));
                Thread.Sleep(2000);
                IJavaScriptExecutor js = (IJavaScriptExecutor)TestSetup.driver;
                bool imgPresent = (bool)js.ExecuteScript("return arguments[0].complete && typeof arguments[0].naturalWidth != 'undefined' && arguments[0].naturalWidth > 0", logoElement);
                if (imgPresent)
                {
                    Console.WriteLine("Image is present");
                    Image = "true"; // Set Image variable to true if image is present
                }
                else
                {
                    Console.WriteLine("Image is not present");
                }
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Image is not present");
            }
        }




    }
}
