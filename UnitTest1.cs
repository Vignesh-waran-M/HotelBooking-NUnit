using OpenQA.Selenium;
using System;
using OpenQA.Selenium.Chrome;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;
using AventStack.ExtentReports;
using System.IO;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework.Interfaces;
using System.Xml;
using HotelBookinNunit.Pages;
using HotelBookinNunit.Rectifiers;
using System.Data;
using OpenQA.Selenium.Firefox;

namespace HotelBookinNunit
{
    public class ExtentReportTests
    {
        public static String xmlpath = "C:\\Users\\vivigneshwaran\\source\\repos\\HotelBookinNunit\\HotelBookinNunit\\DataFile.xml";
        XmlTextReader xtr = new XmlTextReader(xmlpath);
        public static ExtentReports _extent;
        public ExtentTest _test;
        public String TC_Name;
        static string basePath = AppDomain.CurrentDomain.BaseDirectory.Remove(AppDomain.CurrentDomain.BaseDirectory.IndexOf("bin"));

        
        [OneTimeSetUp]
        protected void ExtentStart()
        {
            var path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            var actualPath = path.Substring(0, path.LastIndexOf("bin"));
            var projectPath = new Uri(actualPath).LocalPath;
            Directory.CreateDirectory(projectPath.ToString() + "Reports");

            Console.WriteLine(projectPath.ToString());
            var reportPath = projectPath + "Reports\\Index.html";
            Console.WriteLine(reportPath);
            
            var htmlReporter = new ExtentHtmlReporter(reportPath);
            _extent = new ExtentReports();
            _extent.AttachReporter(htmlReporter);
            htmlReporter.LoadConfig(projectPath + "report-config.xml");
        }
        public IWebDriver driver;
        public WebDriverWait wait;
        public double TotalAmount;
        public string Url;
        ExcelClass exl = new ExcelClass();
        DataTable dt = null;
        
	[OneTimeSetUp]
        public void Setup()
        {
            while (xtr.Read())
            {
                if(xtr.NodeType == XmlNodeType.Element && xtr.Name == "url")
                {
                    Url = xtr.ReadElementContentAsString();
                }
            }
            dt = exl.Input(); //dt will store all the valuees from excel
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            
            driver.Navigate().GoToUrl(Url);
        }
       

        [OneTimeTearDown]
        protected void ExtentClose()
        {            
            _extent.Flush();
        }

        [TearDown]
        public void Cleanup()
        {
            bool passed = TestContext.CurrentContext.Result.Outcome.Status == NUnit.Framework.Interfaces.TestStatus.Passed;
            var exec_status = TestContext.CurrentContext.Result.Outcome.Status;
            var stacktrace = string.IsNullOrEmpty(TestContext.CurrentContext.Result.StackTrace) ? ""
            : string.Format("{0}", TestContext.CurrentContext.Result.StackTrace);
            Status logstatus = Status.Pass;
            String screenShotPath, fileName;
            DateTime time = DateTime.Now;
            fileName = TC_Name + time.ToString("h_mm_ss") + ".png";

            switch (exec_status)
            {
                case TestStatus.Failed:
                    logstatus = Status.Fail;
                    /* The older way of capturing screenshots */
                    screenShotPath = Capture(driver, fileName);
                    /* Capturing Screenshots using built-in methods in ExtentReports 4 */
                    var mediaEntity = CaptureScreenShot(driver, fileName);
                    _test.Log(Status.Fail, "Traditional Snapshot below: " + _test.AddScreenCaptureFromPath("Screenshots\\" + fileName));
                    
                    break;
                case TestStatus.Passed:
                    logstatus = Status.Pass;
                    /* The older way of capturing screenshots */
                    screenShotPath = Capture(driver, fileName);
                    /* Capturing Screenshots using built-in methods in ExtentReports 4 */
                    mediaEntity = CaptureScreenShot(driver, fileName);
                    _test.Log(Status.Pass, "Traditional Snapshot below: " + _test.AddScreenCaptureFromPath("Screenshots\\" + fileName));
                    
                    break;
                case TestStatus.Inconclusive:
                    logstatus = Status.Warning;
                    
                    break;
                case TestStatus.Skipped:
                    logstatus = Status.Skip;
                    
                    break;
                default:
                    
                    break;
            }
            _test.Log(logstatus, "Test: " + TC_Name + " Status:" + logstatus + stacktrace);
            _extent.Flush();

        }

        public static string Capture(IWebDriver driver, String screenShotName)
        {
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            Screenshot screenshot = ts.GetScreenshot();
            var pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            var actualPath = pth.Substring(0, pth.LastIndexOf("bin"));
            var reportPath = new Uri(actualPath).LocalPath;
            Directory.CreateDirectory(reportPath + "Reports\\" + "Screenshots");
            var finalpth = pth.Substring(0, pth.LastIndexOf("bin")) + "Reports\\Screenshots\\" + screenShotName;
            var localpath = new Uri(finalpth).LocalPath;
            screenshot.SaveAsFile(localpath, ScreenshotImageFormat.Png);
            return reportPath;
        }

        public MediaEntityModelProvider CaptureScreenShot(IWebDriver driver, String screenShotName)
        {
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            var screenshot = ts.GetScreenshot().AsBase64EncodedString;

            return MediaEntityBuilder.CreateScreenCaptureFromBase64String(screenshot, screenShotName).Build();
        }

        HomePage homePage = null;
        Cookies cookiePage = null;
        HotelListPage hotelListPage = null;
        SelectHotel selectHotel = null;
        SelectRoomType selectRoomType = null;
        SiteRectifier siteRectifer = null;
        BookRoom bookRoom = null;
        GuestDetails guestDetails = null;
        HeaderAndFooter HandF = null;
        

        [Test,Order(1)]
        public void IamOnTheHolidayInnHomepage()
        {

            String expected_PageTitle = "Holiday Inn® Hotels | Book Family Friendly Hotels Worldwide | Official Site";
            String result_PageTitle;
           

            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;

            _test = _extent.CreateTest(context_name);

            result_PageTitle = driver.Title;
          

            Assert.AreEqual(result_PageTitle, expected_PageTitle);
        }
        
        [Test, Order(2)]
        public void SelectingDestinationAndDate()
        {
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;
            _test = _extent.CreateTest(context_name);
            homePage = new HomePage(driver);
            cookiePage = new Cookies(driver);
            HandF = new HeaderAndFooter(driver);
            cookiePage.ClearCookie();
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            homePage.FillDetails(dt);
        }

        [Test, Order(3)]
        public void IamViewingListOfHotels()
        {
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;

            _test = _extent.CreateTest(context_name);
            hotelListPage = new HotelListPage(driver);
            cookiePage.ClearCookie();
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            hotelListPage.SearchHotel();
        }
	[Test, Order(4)]
        public void SelectingTheHotel()
        {
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;
            _test = _extent.CreateTest(context_name);
            selectHotel = new SelectHotel(driver);
            cookiePage.ClearCookie();
            //HandF = new HeaderAndFooter(driver);
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            selectHotel.HotelSelection();
        }

        [Test,Order(5)]
        public void SelectingRoomTypeAndRate()
        {
            selectRoomType = new SelectRoomType(driver);
            siteRectifer = new SiteRectifier(driver);
            //siteRectifer.Rectify();
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;
            _test = _extent.CreateTest(context_name);
            //HandF = new HeaderAndFooter(driver);
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            selectRoomType.SelectRoomandRate();
            

        }
        [Test,Order(6)]
        public void BookingRoom()
        {
            //siteRectifer.Rectify();
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;
            _test = _extent.CreateTest(context_name);
            bookRoom = new BookRoom(driver);
            //HandF = new HeaderAndFooter(driver);
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            bookRoom.RoomBooking();
        }
        [Test,Order(7)]
        public void GuestInfo()
        {
            String context_name = TestContext.CurrentContext.Test.Name;
            TC_Name = context_name;
            _test = _extent.CreateTest(context_name);
            guestDetails = new GuestDetails(driver);
            //HandF = new HeaderAndFooter(driver);
            HandF.CheckHeader();
            Console.WriteLine("Header elements present");
            HandF.CheckFooter();
            Console.WriteLine("Footer elements present");
            guestDetails.PaymentDetails();
            guestDetails.FillGuestDetails(dt);
        }

    }
}
