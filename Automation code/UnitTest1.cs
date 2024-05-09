using AventStack.ExtentReports;
using AventStack.ExtentReports.Model;
using Bytescout.Spreadsheet;
using OpenQA.Selenium;
using AventStack.ExtentReports.Reporter;
using OpenQA.Selenium.Chrome;
using System.Runtime.Intrinsics.X86;
using static System.Net.Mime.MediaTypeNames;
using OpenQA.Selenium.Support.UI;
using System.Xml.Linq;
using Bytescout.Spreadsheet.COM;
using System;
using System.Globalization;
using System.IO;
using SeleniumExtras.WaitHelpers;

namespace MoFinalProj
{
    [TestClass]
    public class UnitTest1
    {
        // =========================================================================================== 
        // CONFIGURATION
        // ===========================================================================================

        Users user1;
        Users user2;
        Users user3;
        Users user4;


        public static ExtentReports extentReports = new ExtentReports();
        public static ExtentHtmlReporter reporter = new ExtentHtmlReporter("D:\\Files\\QAPractical\\MoFinalProj\\TestReport\\");


        [ClassInitialize]
        public static void InitiateClass(TestContext testContext)
        {
            extentReports.AttachReporter(reporter);
            TestSetup.OpenDriver();
        }

        [ClassCleanup]
        public static void CleanUpClass()
        {
            extentReports.Flush();
            TestSetup.CloseDriver();
        }



        private bool CheckIfAdminSignedIn()
        {
            try
            {
                TestSetup.AdminPage();

                IWebElement expectedElement = TestSetup.driver.FindElement(By.XPath("//h1[text()='ADMIN']"));
                return true;  // Admin is signed up
            }
            catch (NoSuchElementException)
            {
                return false;  // Admin is not signed up
            }
        }


        public void LoginToWebsite()
        {
            bool isSignedUp = CheckIfAdminSignedIn();
            if (!isSignedUp)
            {
                Worksheet LoginSheet = TestSetup.ReadExcel("LoginData");
                user1 = new Users();
                Login Login1 = new Login(user1);
                Login1.StartLogin();
                int j = 2;
                user1.Username = (string)LoginSheet.Cell(j, 0).Value;
                user1.Password = (string)LoginSheet.Cell(j, 1).Value;
                Login1.EnterUsername();
                Login1.EnterPassword();
                Login1.SignInBtn();
            }
        }



        // ===========================================================================================
        // MAIN TEST CASES:
        // ===========================================================================================


        // ------------------------------------------------------------------------------------------- Login Test Case
        public void TestLogin()
        {
            var test = extentReports.CreateTest("Verify Login To Website", Login.LoginString());

            try
            {
                Login Login1 = new Login(user1);
                Login1.StartLogin();
                Login1.EnterUsername();
                Login1.EnterPassword();
                Login1.SignInBtn();

                // Two ways to check
                Thread.Sleep(1000);
                IWebElement ExpectedElement = TestSetup.driver.FindElement(By.XPath("//h1[text()='ADMIN']"));
                TestSetup.CheckElement(ExpectedElement);
                Console.WriteLine("Login successfully");

                string ExpectedResult = "http://localhost:4200/admin/home";
                string ActualResult = TestSetup.driver.Url;
                Assert.AreEqual(ExpectedResult, ActualResult, "actual URL not equal to the expected URL");
                Console.WriteLine(ActualResult);
                test.Pass("Login successfully");

            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Failed");
                test.Fail("Failed successfully :)) ");

                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);

            }
        }

        [TestMethod]
        public void AVerifyLoginToWebsite()
        {
            Worksheet sheet = TestSetup.ReadExcel("LoginData");

            int i = 2;
            user1 = new Users();
            user1.Username = (string)sheet.Cell(i, 0).Value;
            user1.Password = (string)sheet.Cell(i, 1).Value;
            TestLogin();   

        }

        // ------------------------------------------------------------------------------------------- Create Category Test Case

        public void TestCreateCategory(Worksheet sheet, int i)
        {
            var test = extentReports.CreateTest("Verify Created Category", Category.CategoryString());

            TestSetup.AdminPage();

            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));
            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));
            Thread.Sleep(1000);

            int RowCount = TrElements.Count - 1; // Subtract 1 to exclude the header row
            int RowNumber = RowCount; // Initialize with rowCount, which represents the last row number

            // Get the last row in the table
            var LastRow = TrElements[RowCount];
            List<IWebElement> TdElements = new List<IWebElement>(LastRow.FindElements(By.TagName("td")));
            TestSetup.ScrollElement(LastRow);
            TestSetup.CheckElement(LastRow);

            string RowID = TdElements[0].Text.Split(' ')[1];

            string Billname = TdElements[1].Text;
            string Image = "false"; // the default value to false
            string Email = TdElements[3].Text;
            string Location = TdElements[2].Text;
            string Deducted = TdElements[4].Text;

            //  ImagePresentProcess method to update Image variable
            Category.ImagePresentProcess(RowNumber, ref Image);

            if (Billname == (string)sheet.Cell(i, 0).Value && Image == "true"
                && Email == (string)sheet.Cell(i, 2).Value && Location == (string)sheet.Cell(i, 3).Value
                && Deducted == (string)sheet.Cell(i, 4).Value)
            {
                test.Pass("The provided data exists in the table," + " Category ID: " + RowID);
                Console.WriteLine("The provided data exists in the table," + " Category ID: " + RowID);
            }
            else
            {
                if (Billname != (string)sheet.Cell(i, 0).Value)
                {test.Fail("BillName Failed to Verify," + " Category ID: " + RowID);}
                else{test.Pass("BillName Verified Successfully");}

                if (Image != "true")
                {test.Fail("Image Failed to Verify," + " Category ID: " + RowID);}
                else{test.Pass("Image Verified Successfully");}

                if (Email != (string)sheet.Cell(i, 2).Value)
                {test.Fail("Email Failed to Verify," + " Category ID: " + RowID);}
                else{test.Pass("Email Verified Successfully");}

                if (Location != (string)sheet.Cell(i, 3).Value)
                {test.Fail("Location Failed to Verify," + " Category ID: " + RowID);}
                else{test.Pass("Location Verified Successfully");}

                if (Deducted != (string)sheet.Cell(i, 4).Value)
                {test.Fail("Deducted Failed to Verify," + " Category ID: " + RowID);}
                else{test.Pass("Deducted Verified Successfully");}


                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }
        }

        [TestMethod]
        public void BVerifyCreateCategory()
        {
            Worksheet sheet = TestSetup.ReadExcel("CategoryData");

            LoginToWebsite();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user2 = new Users();
                user2.Billname = (string)sheet.Cell(i, 0).Value;
                user2.Image = (string)sheet.Cell(i, 1).Value;
                user2.Email = (string)sheet.Cell(i, 2).Value;
                user2.Location = (string)sheet.Cell(i, 3).Value;
                user2.deducted = (string)sheet.Cell(i, 4).Value;

                Category Category1 = new Category(user2);
                Category1.CreateCategory();
                Category1.EnterBillname();
                Category1.ChooseImage(user2.Image);
                Category1.EnterEmail();
                Category1.EnterLocation();
                Category1.EnterDeducted();
                Category1.CreateBtn();

                TestCreateCategory(sheet, i);
            }

        }

        // ------------------------------------------------------------------------------------------- Update Category Test Case

        public void TestUpdatedCategory(Worksheet sheet, int i, int SelectedRowNumber)
        {
            var test = extentReports.CreateTest("Verify Updated Company", Category.UpdateCategoryString());

            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));
            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));

            var Row = TrElements[SelectedRowNumber];

            List<IWebElement> TdElements = new List<IWebElement>(Row.FindElements(By.TagName("td")));
            TestSetup.ScrollElement(Row);
            TestSetup.CheckElement(Row);

            string RowID = TdElements[0].Text.Split(' ')[1];

            string Billname = TdElements[1].Text;
            string Image = "false"; // the default value to false
            string Email = TdElements[3].Text;
            string Location = TdElements[2].Text;
            string Deducted = TdElements[4].Text;

            //  ImagePresentProcess method to update Image variable
            Category.ImagePresentProcess(SelectedRowNumber, ref Image);

            if (Billname == (string)sheet.Cell(i, 0).Value && Image == "true"
                && Email == (string)sheet.Cell(i, 2).Value && Location == (string)sheet.Cell(i, 3).Value
                && Deducted == (string)sheet.Cell(i, 4).Value)
            {
                test.Pass("The provided data Updated Successfully," + " Category ID: " + RowID);
                Console.WriteLine("The provided data Updated Successfully," + " Category ID: " + RowID);
            }
            else
            {
                if (Billname != (string)sheet.Cell(i, 0).Value)
                { test.Fail("Updated BillName Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated BillName Verified Successfully"); }

                if (Image != "true")
                { test.Fail("Updated Image Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Image Verified Successfully"); }

                if (Email != (string)sheet.Cell(i, 1).Value)
                { test.Fail("Updated Email Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Email Verified Successfully"); }

                if (Location != (string)sheet.Cell(i, 2).Value)
                { test.Fail("Updated Location Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Location Verified Successfully"); }

                if (Deducted != (string)sheet.Cell(i, 3).Value)
                { test.Fail("Updated Deducted Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Deducted Verified Successfully"); }


                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }
        }
        public void UpdateCategory(Worksheet sheet, int i)
        {
            Thread.Sleep(1000);
            TestSetup.AdminPage();

            // Highlight the Table
            IWebElement HighlightTable = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));
            TestSetup.CheckElement(HighlightTable);

            // Get the table by ID - don't want to assume that it's the only table on the page  
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));

            // Get all the rows in the table  
            var Rows = Table.FindElements(By.TagName("tr"));

            // Randomly select a row  
            var Random = new Random();
            var RowNumber = Random.Next(1 , Rows.Count);  // Exclude header row

            var Row = Rows[RowNumber];
            // Find the button within the chosen row and click it  
            var EditBtn = Row.FindElement(By.XPath("./td[6]/button[2]"));
            TestSetup.ScrollElement(EditBtn);
            TestSetup.HighlightElement(EditBtn);
            EditBtn.Click();

            Category Category1 = new Category(user2);
            Category1.ClearUpdateInputField();
            Category1.UpdatBillname();
            Category1.UpdateEmail();
            Category1.UpdatLocation();
            Category1.UpdatDeducted();
            Category1.UpdateImage(user2.Image);
            Category1.UpdatBtn();

            TestUpdatedCategory(sheet, i, RowNumber);
        }

        [TestMethod]
        public void CVerifyUpdateCategory()
        {
            Worksheet sheet = TestSetup.ReadExcel("UpdateCategoryData");

            LoginToWebsite();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user2 = new Users();
                user2.Billname = (string)sheet.Cell(i, 0).Value;
                user2.Email = (string)sheet.Cell(i, 1).Value;
                user2.Location = (string)sheet.Cell(i, 2).Value;
                user2.deducted = (string)sheet.Cell(i, 3).Value;
                user2.Image = (string)sheet.Cell(i, 4).Value;

                UpdateCategory(sheet, i);

            }
        }


        // ------------------------------------------------------------------------------------------- Create Company Test Case

        public void TestCreatedCompany(Worksheet sheet, int i, int SelectedRowNumber)
        {
            var test = extentReports.CreateTest("Verify Created Company", Category.CompanyString());

            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));
            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));

            var Row = TrElements[SelectedRowNumber];
            TestSetup.ScrollElement(Row);

            var CompanyList = Row.FindElements(By.XPath("//td/table/tbody/tr"));

            bool MatchFound = false; // Boolean variable to track if a match is found

            string ID = ""; 

            foreach (var CompanyRow in CompanyList)
            {
                var CompanyData = new List<IWebElement>(CompanyRow.FindElements(By.TagName("td")));
                TestSetup.CheckElement(CompanyRow);

                var BillName = CompanyData[1].Text;
                var Email = CompanyData[2].Text;
                var Location = CompanyData[3].Text;
                if (BillName == (string)sheet.Cell(i, 0).Value && Email == (string)sheet.Cell(i, 1).Value
                    && Location == (string)sheet.Cell(i, 2).Value)
                {
                    MatchFound = true; // Set the boolean variable to true if a match is found

                    string RowID = CompanyData[0].Text.Split(' ')[0];
                    ID = RowID;
                }
            }

            if (MatchFound)
            {
                test.Pass("The provided data exists in the selected Category, Company ID: " + ID);
                Console.WriteLine("The provided data exists in the selected Category, Company ID: " + ID);
            }
            else
            {
                test.Fail("The provided data does not exist in the Category");
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }
        }

        public void CreateCompany(Worksheet sheet, int i)
        {
            Thread.Sleep(1000);
            TestSetup.AdminPage();

            // Get the table by ID - don't want to assume that it's the only table on the page  
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));

            // Get all the rows in the table  
            var Rows = Table.FindElements(By.TagName("tr"));
            TestSetup.CheckElement(Table);

            // Randomly select a row  
            var Random = new Random();
            var RowNumber = Random.Next(1, Rows.Count);  // Exclude header row
            var Row = Rows[RowNumber];

            // Find the button within the chosen row and click it  
            var button = Row.FindElement(By.XPath("./td[6]/button[1]"));
            TestSetup.ScrollElement(button);
            TestSetup.HighlightElement(button);
            button.Click();

            Category Category1 = new Category(user2);
            Category1.EnterCompanyBillname();
            Category1.EnterEmail();
            Category1.EnterLocation();
            Category1.CreateCompanyBtn();
            Thread.Sleep(1000);

            try
            {
                var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
                TestSetup.ScrollElement(ShowBtn);
                TestSetup.HighlightElement(ShowBtn);
                ShowBtn.Click();
                Thread.Sleep(1000);
                TestCreatedCompany(sheet, i, RowNumber);
            }
            catch (StaleElementReferenceException)
            {
                // Re-find the elements  
                Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));
                Rows = Table.FindElements(By.TagName("tr"));
                Row = Rows[RowNumber];
                var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
                TestSetup.ScrollElement(ShowBtn);
                TestSetup.HighlightElement(ShowBtn);
                ShowBtn.Click();
                Thread.Sleep(1000);

                TestCreatedCompany(sheet, i, RowNumber);
            }

        }

        [TestMethod]
        public void DVerifyCreateCompany()
        {
            Worksheet sheet = TestSetup.ReadExcel("CompanyData");

            LoginToWebsite();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user2 = new Users();
                user2.CompanyBillname = (string)sheet.Cell(i, 0).Value;
                user2.Email = (string)sheet.Cell(i, 1).Value;
                user2.Location = (string)sheet.Cell(i, 2).Value;

                CreateCompany(sheet, i);
            }
        }

        // ------------------------------------------------------------------------------------------- Search By Category Name Test case

        [TestMethod]
        public void EVerifyCategorySearchByName()
        {
            LoginToWebsite();

            for (int i = 0; i < 2; i++)
            {
                var test = extentReports.CreateTest("Verify Search By Category Name", Report.ReportNameString());

                Report Report1 = new Report(user4);
                Report1.StartReport();
                Thread.Sleep(2000);
                IWebElement ReportDropdownList = TestSetup.driver.FindElement(By.XPath("//select[@formcontrolname='Billercategoryname']"));
                TestSetup.HighlightElement(ReportDropdownList);
                ReportDropdownList.SendKeys(Keys.Right);

                // Get all the options within the dropdown
                List<IWebElement> Options = new List<IWebElement>(ReportDropdownList.FindElements(By.TagName("option")));

                // Generate a random index within the range of available options
                Random Random = new Random();
                int RandomIndex = Random.Next(1,Options.Count);

                Console.WriteLine(RandomIndex);
                // Select the option at the randomly generated index
                Options[RandomIndex].Click();
                Thread.Sleep(3000);


                // Locate Category Table
                IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-adminreport/div/div/div[2]/div/table"));

                // Locate the selected option index and extract it text
                IWebElement SelectedOption = Options[RandomIndex];
                string SelectedOptionText = SelectedOption.Text;
                var Rows = Table.FindElements(By.TagName("tr"));

                foreach (var row in Rows.Skip(1))

                {

                    try
                    {
                        Console.WriteLine("Date1");
                        var Name = row.FindElement(By.XPath(".//td[1]"));
                        TestSetup.CheckElement(Name);
                        if (SelectedOptionText == Name.Text) { test.Pass("Search by Category Name Verified Successfully");break; }


                    }

                    catch
                    {
                        test.Fail("Search by Category Name Failed to Verify");
                        string fullPath = TestSetup.TakeScreenShot();
                        test.AddScreenCaptureFromPath(fullPath);
                    }

                }


                Report1.ClearDropDownList();
            }

        }


        // ------------------------------------------------------------------------------------------- Search By Category Date Test case

        public void TestPaymentDate(Worksheet sheet, int i)
        {
            var test = extentReports.CreateTest("Verify Search By Category Date", Report.ReportDateString());

            Thread.Sleep(2000);
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-detels/div/div/div/div/table"));
            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));
            TestSetup.CheckElement(Table);

            Thread.Sleep(2000);

            int RowCount = TrElements.Count;

            Random Random = new Random();
            int RandomIndex = Random.Next(1, RowCount);

            var DateElement = TestSetup.driver.FindElement(By.XPath($"/html/body/app-root/app-detels/div/div/div/div/table/tbody/tr[{RandomIndex}]/td[4]"));
            TestSetup.ScrollElement(DateElement);
            TestSetup.CheckElement(DateElement);
            Thread.Sleep(1000);
            string RowDate = DateElement.Text;

            // Interval dates from the Excel sheet
            string FromDate = (string)sheet.Cell(i, 0).Value;
            string ToDate = (string)sheet.Cell(i, 1).Value;

            // Convert the extracted dates to DateTime objects
            DateTime FromDateTime = DateTime.ParseExact(FromDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            DateTime ToDateTime = DateTime.ParseExact(ToDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            DateTime RowDateTime = DateTime.ParseExact(RowDate, "MMM d, yyyy", CultureInfo.InvariantCulture);

            // Check if the row date is within the interval
            if (RowDateTime >= FromDateTime && RowDateTime <= ToDateTime)
            {
                Console.WriteLine("Date: " + RowDate + " is within the interval: " + FromDate + " - " + ToDate);
                test.Pass("Date: " + RowDate + " is within the interval: " + FromDate + " - " + ToDate);
            }
            else
            {
                Console.WriteLine("Date: " + RowDate + " is not within the interval: " + FromDate + " - " + ToDate);
                test.Fail("Date: " + RowDate + " is not within the interval: " + FromDate + " - " + ToDate);
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }
        }

        public void TestCategoryDate(Worksheet sheet, int i)
        {
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-adminreport/div/div/div[2]/div/table"));
            var Rows = Table.FindElements(By.TagName("tr"));
            TestSetup.CheckElement(Table);


            if (Rows.Count > 2)
            {
                Random Random = new Random();
                int RandomIndex = Random.Next(1, Rows.Count - 1);
                Console.WriteLine(Rows.Count);
                Console.WriteLine(RandomIndex);

                IWebElement Button = TestSetup.driver.FindElement(By.XPath($"/html/body/app-root/app-adminreport/div/div/div[2]/div/table/tbody/tr[{RandomIndex}]/td[4]/button"));
                TestSetup.HighlightElement(Button);
                Button.Click();

                TestPaymentDate(sheet, i);
                Console.WriteLine("All done.");
            }
            else
            {
                var test = extentReports.CreateTest("Verify Search By Category Date", Report.ReportDateString());
                Console.WriteLine("There is no row to check.");
                test.Pass("There is no row to check.");
            }
        }

        [TestMethod]
        public void FVerifyCategorySearchByDate()
        {
            Worksheet sheet = TestSetup.ReadExcel("ReportData");

            LoginToWebsite();

            Report Report1 = new Report(user4);
            Report1.StartReport();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user4 = new Users();
                user4.ReportDateFrom = (string)sheet.Cell(i, 0).Value;
                user4.ReportDateTo = (string)sheet.Cell(i, 1).Value;

                Report1.ReportDropdownList();
                Report1.ReportDate(user4.ReportDateFrom, user4.ReportDateTo);

                TestCategoryDate(sheet, i);

                TestSetup.ReportPage();

            }
        }


        // ------------------------------------------------------------------------------------------- Payment Count Test case

        public void TestPyamentCount(int RowsNumber)
        {
            var test = extentReports.CreateTest("Verify Payment Count", Report.ReportCountString());

            Thread.Sleep(2000);
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-detels/div/div/div/div/table"));
            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));
            TestSetup.CheckElement(Table);

            Thread.Sleep(2000);

            int CountPay = RowsNumber;
            int CountRow = TrElements.Count;
            int RowCount = CountRow - 1;

            try
            {
                Assert.AreEqual(CountPay, RowCount);
                test.Pass("CountPay and Number of payment is equal");
            }
            catch (NoSuchElementException)
            {
                Assert.AreEqual(CountPay, RowCount);
                test.Pass("CountPay and Number of payment is equal");
            }

            catch
            {
                test.Fail("CountPay: " + CountPay + " and Number of payment: " + RowCount + " is not equal");
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }

        }

        public void PyamentCount(Worksheet sheet, int i)
        {
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-adminreport/div/div/div[2]/div/table"));
            var Rows = Table.FindElements(By.TagName("tr"));
            TestSetup.CheckElement(Table);


            if (Rows.Count > 2)
            {
                Random Random = new Random();
                int RandomIndex = Random.Next(1, Rows.Count - 1);
                Console.WriteLine(Rows.Count);
                Console.WriteLine(RandomIndex);

                IWebElement Count = TestSetup.driver.FindElement(By.XPath($"/html/body/app-root/app-adminreport/div/div/div[2]/div/table/tbody/tr[{RandomIndex}]/td[2]"));
                TestSetup.CheckElement(Count);
                int CountPay = int.Parse(Count.Text);
                Console.WriteLine(CountPay);

                IWebElement HistoryBtn = TestSetup.driver.FindElement(By.XPath($"/html/body/app-root/app-adminreport/div/div/div[2]/div/table/tbody/tr[{RandomIndex}]/td[4]/button"));
                TestSetup.HighlightElement(HistoryBtn);
                HistoryBtn.Click();

                TestPyamentCount(CountPay);
                Console.WriteLine("All done.");
            }
            else
            {
                var test = extentReports.CreateTest("Verify Payment Count", Report.ReportCountString());
                Console.WriteLine("There is no row to check.");
                test.Pass("There is no row to check.");
            }
        }

        [TestMethod]
        public void GVerifyPyamentCount()
        {
            Worksheet sheet = TestSetup.ReadExcel("ReportData");

            LoginToWebsite();

            Report Report1 = new Report(user4);
            Report1.StartReport();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user4 = new Users();
                user4.ReportDateFrom = (string)sheet.Cell(i, 0).Value;
                user4.ReportDateTo = (string)sheet.Cell(i, 1).Value;

                Report1.ReportDropdownList();
                Report1.ReportDate(user4.ReportDateFrom, user4.ReportDateTo);

                PyamentCount(sheet , i);

                TestSetup.ReportPage();
            }

        }


        // ------------------------------------------------------------------------------------------- Change FullName Test Case

        public void TestUpdateFullName()
        {
            var test = extentReports.CreateTest("Verify Updating fullName", UpdateName.ProfileString());

            LoginToWebsite();

            UpdateName UpdateName1 = new UpdateName(user3);
            UpdateName1.StartProfile();
            UpdateName1.ClearInputField();
            UpdateName1.UpdateFullName();
            UpdateName1.EnterPassword();
            UpdateName1.UpdateBtn();

            try
            {
                Thread.Sleep(1000);
                IWebElement VerfiyFulltName = TestSetup.driver.FindElement(By.XPath($"//input[@formcontrolname='FullName' and @ng-reflect-model='{user3.FullName}']"));
                TestSetup.CheckElement(VerfiyFulltName);
                Console.WriteLine("Verify Updated FullName successfully");
                test.Pass("FullName updated successfully");
            }
            catch
            {
                Console.WriteLine("Failed to verify Updated FullName");
                test.Fail("Failed to update FullName");

                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }

        }

        [TestMethod]
        public void HVerifyUpdateFullName()
        {
            Worksheet sheet = TestSetup.ReadExcel("UpdateNameData");

            int i = 2;
            user3 = new Users();
            user3.FullName = (string)sheet.Cell(i, 0).Value;
            user3.Password = (string)sheet.Cell(i, 1).Value;
            TestUpdateFullName();
        }


        // ===========================================================================================
        // EXTRA TEST CASES:
        // ===========================================================================================


        // ------------------------------------------------------------------------------------------- Create Company When "Show" Button Pressed Test Case
        public void TestCreatedCompanyWithShowButton(Worksheet sheet, int i, ExtentTest test, int SelectedRow)
        {
            IWebElement Table = TestSetup.driver.FindElement(By.XPath($"//td/table/tbody[{SelectedRow + 1}]/tr"));

            var CompanyData = new List<IWebElement>(Table.FindElements(By.TagName("td")));
            TestSetup.CheckElement(Table);

            string RowID = CompanyData[0].Text.Split(' ')[0];

            var BillName = CompanyData[1].Text;
            var Email = CompanyData[2].Text;
            var Location = CompanyData[3].Text;


            if (BillName == (string)sheet.Cell(i, 0).Value && Email == (string)sheet.Cell(i, 1).Value
                && Location == (string)sheet.Cell(i, 2).Value)
            {
                test.Pass("The provided data exists in the selected Category," + " Company ID: " + RowID);
                Console.WriteLine("The provided data exists in the selected Category," + " Company ID: " + RowID);
            }
            else
            {
                test.Fail("The provided data does not exists in the selected Category");
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);

            }

        }

        public void TestCreateCompanyWithShowButton(Worksheet sheet, int i, int SelectedRow, List<string> CompanyIDsBefore)
        {
            Thread.Sleep(2000);
            TestSetup.AdminPage();
            ExtentTest test = extentReports.CreateTest("Verify Created Company With Show Button Pressed", Category.CompanyShowButtonString());

            WebDriverWait wait = new WebDriverWait(TestSetup.driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementExists(By.XPath("/html/body/app-root/app-home/div/div/div/table")));

            // Get the table by ID - don't want to assume that it's the only table on the page  
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));

            wait.Until(ExpectedConditions.ElementExists(By.TagName("tr")));

            // Get all the rows in the table  
            var Rows = Table.FindElements(By.TagName("tr"));
            Console.WriteLine("sec SelectedRow: " + SelectedRow);
            Console.WriteLine("sec count: " + Rows.Count);
            var Row = Rows[SelectedRow];

            var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
            TestSetup.ScrollElement(ShowBtn);
            TestSetup.HighlightElement(ShowBtn);
            ShowBtn.Click();

            var CompanyList = Row.FindElements(By.XPath("//td/table/tbody/tr"));
            TestSetup.ScrollElement(Row);

            // Compare companyIDsBefore with companyIDsAfter
            List<string> CompanyIDsAfter = new List<string>(); // List to store company IDs
            bool ExistingIDs = false;

            foreach (var CompanyRow in CompanyList)
            {

                var CompanyData = new List<IWebElement>(CompanyRow.FindElements(By.TagName("td")));
                var CompanyID = CompanyData[0].Text;

                if (!string.IsNullOrEmpty(CompanyID))
                {
                    CompanyIDsAfter.Add(CompanyID); // Add company ID to the list
                    ExistingIDs = true;
                }
            }

            if (!ExistingIDs)
            {
                CompanyIDsAfter.Add("-1"); // Add -1 to the list if no existing IDs are found
            }

            Console.WriteLine(string.Join(", ", CompanyIDsAfter));

            // Compare companyIDsBefore and companyIDsAfter
            if (CompanyIDsBefore[0] == "-1" && CompanyIDsAfter[0] == "-1")
            {
                Console.WriteLine("No rows (Companies) were added");
                test.Fail("The provided data does not get added when show button pressed");
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }

            else
            {
                // Compare companyIDsBefore with companyIDsAfter
                List<string> differentItems = CompanyIDsAfter.Except(CompanyIDsBefore).ToList();
                Console.WriteLine(differentItems);
                // Print the differences
                Console.WriteLine("Differences between companyIDsBefore and companyIDsAfter: " + string.Join(", ", differentItems));

                int SelectedItem = 0;
                for (int j = 0; j < CompanyIDsAfter.Count; j++)
                {

                    if (CompanyIDsAfter[j] == differentItems[0])
                    {
                        SelectedItem = j;
                        break;
                    }
                }
                Console.WriteLine("index: " + SelectedItem);

                TestCreatedCompanyWithShowButton(sheet, i, test, SelectedItem);
            }

        }
        public void CreateCompanyWithShowButton(Worksheet sheet, int i)
        {
            Thread.Sleep(1000);
            TestSetup.AdminPage();

            WebDriverWait wait = new WebDriverWait(TestSetup.driver, TimeSpan.FromSeconds(20));
            wait.Until(ExpectedConditions.ElementExists(By.XPath("/html/body/app-root/app-home/div/div/div/table")));

            // Get the table by ID - don't want to assume that it's the only table on the page
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));


            // Get all the rows in the table
            var Rows = Table.FindElements(By.TagName("tr"));
            TestSetup.CheckElement(Table);

            // Randomly select a row
            var Random = new Random();
            var SelectedRow = Random.Next(1, Rows.Count); // Exclude header row
            Console.WriteLine("first SelectedRow: " + SelectedRow);
            Console.WriteLine("first count: " + Rows.Count);
            var Row = Rows[SelectedRow];

            // Wait for Element

            wait.Until(ExpectedConditions.ElementExists(By.XPath(".//td[1]/button[1]")));

            var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
            TestSetup.ScrollElement(ShowBtn);
            TestSetup.HighlightElement(ShowBtn);
            ShowBtn.Click();
            Thread.Sleep(1000);

            var CompanyList = Row.FindElements(By.XPath("//td/table/tbody/tr"));
            TestSetup.ScrollElement(Row);
            Console.WriteLine(" count: " + CompanyList.Count);

            List<string> CompanyIDsBefore = new List<string>(); // List to store company IDs
            bool ExistingIDs = false;

            foreach (var CompanyRow in CompanyList)
            {
                var companyData = new List<IWebElement>(CompanyRow.FindElements(By.TagName("td")));
                TestSetup.CheckElement(CompanyRow);

                var companyID = companyData[0].Text;
                if (!string.IsNullOrEmpty(companyID))
                {
                    CompanyIDsBefore.Add(companyID); // Add company ID to the list
                    ExistingIDs = true;
                }
            }

            if (!ExistingIDs)
            {
                CompanyIDsBefore.Add("-1"); // Add -1 to the list if No existing IDs are found
            }

            Console.WriteLine(string.Join(", ", CompanyIDsBefore));

            // Find the button within the chosen row and click it
            var button = Row.FindElement(By.XPath("./td[6]/button[1]"));
            TestSetup.ScrollElement(button);
            TestSetup.HighlightElement(button);
            button.Click();

            Category Category1 = new Category(user2);
            Category1.EnterCompanyBillname();
            Category1.EnterEmail();
            Category1.EnterLocation();
            Category1.CreateCompanyBtn();

            Console.WriteLine("row num: " + SelectedRow);

            TestCreateCompanyWithShowButton(sheet, i, SelectedRow, CompanyIDsBefore);
        }

        [TestMethod]
        public void YVerifyCreateCompanyWithShowButton()
        {
            Worksheet sheet = TestSetup.ReadExcel("CompanyData");

            LoginToWebsite();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user2 = new Users();
                user2.CompanyBillname = (string)sheet.Cell(i, 0).Value;
                user2.Email = (string)sheet.Cell(i, 1).Value;
                user2.Location = (string)sheet.Cell(i, 2).Value;

                CreateCompanyWithShowButton(sheet, i);
            }
        }





        // ------------------------------------------------------------------------------------------- Update Category When "Show" Button Pressed Test Case

        public void TestUpdatedCategoryWithShowButton(Worksheet sheet, int i, int SelectedRowNumber)
        {
            var test = extentReports.CreateTest("Verify Updated Category when show button Pressed", Category.UpdateCategoryShowButtonString());

            WebDriverWait wait = new WebDriverWait(TestSetup.driver, TimeSpan.FromSeconds(10));
            IWebElement Table = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/app-root/app-home/div/div/div/table")));

            List<IWebElement> TrElements = new List<IWebElement>(Table.FindElements(By.TagName("tr")));
            var Row = TrElements[SelectedRowNumber];
            TestSetup.ScrollElement(Row);
            TestSetup.CheckElement(Row);

            List<IWebElement> TdElements = new List<IWebElement>(Row.FindElements(By.TagName("td")));

            string RowID = TdElements[0].Text.Split(' ')[1];

            string Billname = TdElements[1].Text;
            string Image = "false"; // the default value to false
            string Email = TdElements[3].Text;
            string Location = TdElements[2].Text;
            string Deducted = TdElements[4].Text;

            //  ImagePresentProcess method to update Image variable
            Category.ImagePresentProcess(SelectedRowNumber, ref Image);

            if (Billname == (string)sheet.Cell(i, 0).Value && Image == "true"
                && Email == (string)sheet.Cell(i, 1).Value && Location == (string)sheet.Cell(i, 2).Value
                && Deducted == (string)sheet.Cell(i, 3).Value)
            {

                test.Pass("The provided data Updated Successfully," + " Category ID: " + RowID);
                Console.WriteLine("The provided data Updated Successfully," + " Category ID: " + RowID);
            }
            else
            {
                if (Billname != (string)sheet.Cell(i, 0).Value)
                { test.Fail("Updated BillName Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated BillName Verified Successfully"); }

                if (Image != "true")
                { test.Fail("Updated Image Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Image Verified Successfully"); }

                if (Email != (string)sheet.Cell(i, 1).Value)
                { test.Fail("Updated Email Failed to Verify" + " Category ID: " + RowID); }
                else { test.Pass("Updated Email Verified Successfully"); }

                if (Location != (string)sheet.Cell(i, 2).Value)
                { test.Fail("Updated Location Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Location Verified Successfully"); }

                if (Deducted != (string)sheet.Cell(i, 3).Value)
                { test.Fail("Updated Deducted Failed to Verify," + " Category ID: " + RowID); }
                else { test.Pass("Updated Deducted Verified Successfully"); }

                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }

            try
            {
                var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
                TestSetup.ScrollElement(ShowBtn);
                TestSetup.CheckElement(ShowBtn);
                ShowBtn.Click();
                test.Pass("Category Update Verified successfully");
            }
            catch
            {
                Console.WriteLine("Category Update Failed to Verify");
                test.Fail("Category Update Failed to Verified");
                string fullPath = TestSetup.TakeScreenShot();
                test.AddScreenCaptureFromPath(fullPath);
            }

        }

        public void UpdateCategoryWithShowButton(Worksheet sheet, int i)
        {
            Thread.Sleep(1000);
            TestSetup.AdminPage();

            // Get the table by ID - don't want to assume that it's the only table on the page 
            IWebElement Table = TestSetup.driver.FindElement(By.XPath("/html/body/app-root/app-home/div/div/div/table"));

            // Get all the rows in the table  
            var Rows = Table.FindElements(By.TagName("tr"));
            TestSetup.CheckElement(Table);

            // Randomly select a row  
            var Random = new Random();
            var RowNumber = Random.Next(1, Rows.Count); // Exclude header row  
            var Row = Rows[RowNumber];

            // Find the button within the chosen row and click it
            var ShowBtn = Row.FindElement(By.XPath(".//td[1]/button[1]"));
            TestSetup.ScrollElement(ShowBtn);
            TestSetup.HighlightElement(ShowBtn);
            ShowBtn.Click();
            Thread.Sleep(1000);

            var EditBtn = Row.FindElement(By.XPath("./td[6]/button[2]"));
            TestSetup.ScrollElement(EditBtn);
            TestSetup.HighlightElement(EditBtn);
            EditBtn.Click();

            Category Category1 = new Category(user2);
            Category1.ClearUpdateInputField();
            Category1.UpdatBillname();
            Category1.UpdateEmail();
            Category1.UpdatLocation();
            Category1.UpdatDeducted();
            Category1.UpdateImage(user2.Image);
            Category1.UpdatBtn();

            TestUpdatedCategoryWithShowButton(sheet, i, RowNumber);
        }


        [TestMethod]
        public void ZVerifyUpdateCategoryWithShowButton()
        {
            Worksheet sheet = TestSetup.ReadExcel("UpdateCategoryData");

            LoginToWebsite();

            for (int i = 2; i <= sheet.NotEmptyRowMax; i++)
            {
                user2 = new Users();
                user2.Billname = (string)sheet.Cell(i, 0).Value;
                user2.Email = (string)sheet.Cell(i, 1).Value;
                user2.Location = (string)sheet.Cell(i, 2).Value;
                user2.deducted = (string)sheet.Cell(i, 3).Value;
                user2.Image = (string)sheet.Cell(i, 4).Value;

                UpdateCategoryWithShowButton(sheet, i);
            }
        }





    }
}