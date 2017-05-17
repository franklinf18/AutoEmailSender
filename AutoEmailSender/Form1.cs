using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;

using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;

using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.PhantomJS;
using System.Diagnostics;

namespace AutoEmailSender
{
    public partial class FormAutoEmailSender : Form
    {

        public IWebDriver driver { get; private set; }
        public double completionPercentage { get; private set; }
        public bool addPasswordInfoSucc { get; set; }
        public string webConfigPath { get; set; }
        public string[] managerIntranetIds { get; set; }
        public int iterationCount { get; set; }
        public string emailSender { get; set; }
        public int escalationCounter { get; set; }
        public string screenshotPath { get; set; }
        public string toList { get; set; }
        public string excelFilePath { get; set; }
        public string errorFilePath { get; set; }
        public string dashboardLink { get; set; }



        public FormAutoEmailSender()
        {
            InitializeComponent();
            //TestEncryption();
            //SendEmailINotes();
            SetProperties();
            UpdateCredentialsFromWebConfig();
            CaptureScreenshot();
            //CaptureScreenshotITSec();
            //SendEmail();
            SendEmailINotes();

        }




        public void SendEmail()
        {
            try
            {
                //string path = System.AppDomain.CurrentDomain.BaseDirectory;
                //string[] folders = System.IO.Directory.GetDirectories(path, "*.default", System.IO.SearchOption.AllDirectories);

                //string pathFinal = path + "geckodriver.exe";



                //Environment.SetEnvironmentVariable("webdriver.gecko.driver", pathFinal);
                //FirefoxOptions fOptions = new FirefoxOptions();


                //string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Roaming\Mozilla\Firefox\Profiles";
                //string[] folders = System.IO.Directory.GetDirectories(path, "*.default", System.IO.SearchOption.AllDirectories);

                //string pathFinal = folders[0].ToString();
                //FirefoxProfile ffprofile = new FirefoxProfile(path);
                //IWebDriver driver = new FirefoxDriver(ffprofile);
                //driver = new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe"), new FirefoxProfile(pathFinal));
                //driver = new PhantomJSDriver();

                driver = new InternetExplorerDriver();
                driver.Manage().Window.Maximize();
                string password = "";

                driver.Navigate().GoToUrl("https://mail.notes.na.collabserv.com/sequoia/index.html");
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("Intranet_ID")));

                //IList<IWebElement> usernameInput = driver.FindElements(By.TagName("input"));

                //foreach (var element in usernameInput)
                //{
                //    if (element.GetAttribute("id") == "username")
                //    {
                //        element.SendKeys("frfernan@in.ibm.com");
                //        driver.FindElement(By.Id("continue")).Click();
                //        break;
                //    }

                //}


                //if (Properties.Settings.Default["frfernan@in.ibm.com"] !=null)
                //{
                //    string encryptedPassword = Properties.Settings.Default["frfernan@in.ibm.com"].ToString();
                //    //Decryption
                //    password = DecryptString(encryptedPassword);

                emailSender = System.Configuration.ConfigurationManager.AppSettings["emailsender"];

                if (System.Configuration.ConfigurationManager.AppSettings[emailSender] != null)
                {
                    //MessageBox.Show("Entered Loop");
                    string encryptedPassword = Convert.ToString(ConfigurationManager.AppSettings[emailSender]);
                    if (encryptedPassword != "" || encryptedPassword != null)
                    {
                        //MessageBox.Show("Entered 2nd Loop "+ encryptedPassword);

                    }
                    //Decryption
                    //password = DecryptString(encryptedPassword);
                    password = AesCryp.Decrypt(encryptedPassword);

                    //MessageBox.Show("Entered 2nd Loop " + password);


                    //driver.FindElement(By.Id("Intranet_ID")).Clear();
                    driver.FindElement(By.XPath("/html/body/div/div[2]/div/div[2]/div[1]/div/div/div/form/div/fieldset/p[1]/span/input")).Clear();
                    driver.FindElement(By.Id("Intranet_ID")).SendKeys(emailSender);
                    driver.FindElement(By.Id("password")).Clear();
                    driver.FindElement(By.Id("password")).SendKeys(password);
                    driver.FindElement(By.Name("ibm-submit")).Click();
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=lconn_gadget_lifecycle_publisher | ]]
                    //driver.FindElement(By.LinkText("Mail")).Click();
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=lconn_gadget_lifecycle_publisher | ]]
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=lconn_gadget_lifecycle_publisher | ]]
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=lconn_gadget_lifecycle_publisher | ]]
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]

                    Console.WriteLine("Logged In");
                    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("button.compose-button")));


                    driver.FindElement(By.CssSelector("button.compose-button")).Click();
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]
                    //driver.FindElement(By.Id("dijit__TemplatedMixin_24")).Clear();
                    //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60)); //you can play with the time integer  to wait for longer.
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div/div[1]/div[1]")));

                    //driver.FindElement(By.Id("dijit__TemplatedMixin_24")).Click();
                    //driver.FindElement(By.Id("dijit__TemplatedMixin_24")).SendKeys("frfernan@in.ibm.com");

                    OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver); // define Actions object


                    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div/div[1]/div[1]")).Click();
                    //driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[7]/div[1]/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div/div[1]/div[1]")).SendKeys("frfernan@in.ibm.com; chandanb@in.ibm.com");
                    //action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys("frfernan@in.ibm.com;").Perform();
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]

                    int weekNo = GetIso8601WeekOfYear(DateTime.Now);

                    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[1]/input")).Clear();
                    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[1]/input")).SendKeys("ILC and IT Security Diary Status -- Week " + weekNo);   //Change this LOC once ITSec tracking is completed
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]
                    //driver.FindElement(By.Id("uniqName_85_0")).Click();
                    driver.SwitchTo().Frame(0); // Switch to Body of Email
                    driver.FindElement(By.XPath("/html/body/div[1]")).Click();
                    driver.FindElement(By.XPath("/html/body/div[1]")).SendKeys("Team, \n\nPlease find the ILC and IT Security Diary Status for Week " + weekNo + "\n\n");  //Change this LOC once ITSec tracking is completed
                    driver.FindElement(By.XPath("/html/body/div[1]")).SendKeys("ILC Status: \n\n");  //Change this LOC once ITSec tracking is completed

                    driver.SwitchTo().DefaultContent();

                    PasteImage();

                    //Code to paste ITSec screenshot
                    driver.SwitchTo().DefaultContent();

                    driver.SwitchTo().Frame(0); // Switch to Body of Email
                    driver.FindElement(By.XPath("/html/body/div[1]")).Click();

                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Perform();


                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();
                    action.SendKeys("IT Security Status:");
                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                    driver.SwitchTo().DefaultContent();
                    PasteImageITSec();

                    driver.SwitchTo().DefaultContent();

                    //driver.FindElement(By.XPath("/html/body/div[1]")).SendKeys(OpenQA.Selenium.Keys.Enter);


                    driver.SwitchTo().DefaultContent();

                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    //action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    //action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                    Thread.Sleep(5000);
                    driver.Close();
                    driver.Dispose();
                    // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]
                    //driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[7]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[6]/button[1]")).Click();
                    //bool isSelected = driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[7]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[6]/button[1]")).Selected;
                    //if (isSelected)
                    //{
                    //    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[7]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[6]/button[1]")).Click();

                    //}
                }

                else
                {
                    MessageBox.Show("Data not found in config file. Please register");
                }



            }
            catch (Exception ex)
            {
                this.LogError(ex);
                //throw;
                if (iterationCount <= 3)
                {
                    iterationCount++;
                    RestartEmailSender();

                }
                else
                {
                    driver.Close();
                    driver.Dispose();
                }
            }


        }

        public void SendEmailINotes()
        {
            try
            {
                CloseAllBrowserWindows();
                driver = new InternetExplorerDriver();

                driver.Manage().Window.Maximize();
                string password = "";
                int weekNo = GetIso8601WeekOfYear(DateTime.Now);

                string baseURL = "https://mail.notes.na.collabserv.com/";
                driver.Navigate().GoToUrl(baseURL + "/livemail/iNotes/Mail/?OpenDocument&noredir=1");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("Intranet_ID")));

                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver); // define Actions object

                //Below lines commented as properties are set in SetProperties method
                //emailSender = System.Configuration.ConfigurationManager.AppSettings["emailsender"];
                //screenshotPath = System.Configuration.ConfigurationManager.AppSettings["screenshotPath"];
                //toList = System.Configuration.ConfigurationManager.AppSettings["toList"];

                if (System.Configuration.ConfigurationManager.AppSettings[emailSender] != null)
                {
                    //MessageBox.Show("Entered Loop");
                    string encryptedPassword = Convert.ToString(ConfigurationManager.AppSettings[emailSender]);
                    if (encryptedPassword != "" || encryptedPassword != null)
                    {
                        //MessageBox.Show("Entered 2nd Loop "+ encryptedPassword);

                    }
                    //Decryption
                    //password = DecryptString(encryptedPassword);
                    password = AesCryp.Decrypt(encryptedPassword);

                    //MessageBox.Show("Entered 2nd Loop " + password);


                    //driver.FindElement(By.Id("Intranet_ID")).Clear();
                    driver.FindElement(By.XPath("/html/body/div/div[2]/div/div[2]/div[1]/div/div/div/form/div/fieldset/p[1]/span/input")).Clear();
                    driver.FindElement(By.Id("Intranet_ID")).SendKeys(emailSender);
                    driver.FindElement(By.Id("password")).Clear();
                    driver.FindElement(By.Id("password")).SendKeys(password);
                    driver.FindElement(By.Name("ibm-submit")).Click();

                    // ERROR: Caught exception [ERROR: Unsupported command [selectFrame | s_MainFrame | ]]
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("s_MainFrame")));

                    driver.SwitchTo().Frame("s_MainFrame");

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("e-actions-mailview-inbox-new-text")));

                    driver.FindElement(By.Id("e-actions-mailview-inbox-new-text")).Click();
                    //driver.FindElement(By.Id("e-actions-mailview-inbox-new-text")).Click();

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("e-$new-0-sendto")));

                    driver.FindElement(By.Id("e-$new-0-sendto")).Clear();
                    driver.FindElement(By.Id("e-$new-0-sendto")).SendKeys(toList);
                    driver.FindElement(By.Id("e-$new-0-subject")).Clear();
                    driver.FindElement(By.Id("e-$new-0-subject")).SendKeys("ILC Status -- Week " + weekNo);
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                    action.SendKeys("Team, \n\nPlease find the ILC Status for Week " + weekNo + "\n\n").Perform();
                    action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();
                    //action.SendKeys("ILC Status: \n\n").Perform();

                    //driver.SwitchTo().Frame("e-$new-0-bodyrich-editorframe");
                    //driver.FindElement(By.XPath("/html/body/span/div/div/div/div/div/div/div[1]/font/div/div/div/font/div[1]")).Click();
                    //driver.SwitchTo().ParentFrame();

                    //Image attachment code
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");
                    string latestScreenshotILC = RetrieveLatestScreenshot(screenshotPath);
                    //driver.SwitchTo().Window(driver.WindowHandles[0]);

                    // validate if latestScreenshotILC is not empty =========================

                    if (latestScreenshotILC != "" || latestScreenshotILC != null)
                    {
                        latestScreenshotILC = RetrieveLatestScreenshot(screenshotPath);
                    }

                    else
                    {
                        throw new Exception("Unable to get ILC Screenshot path");
                    }
                    // ======================================================================


                    driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[1]/div/div[5]/div/form/div/div/div[3]/div/div[1]/table/tbody/tr/td[4]/img[18]")).Click();
                    //action.SendKeys(OpenQA.Selenium.Keys.Tab);
                    //action.SendKeys(OpenQA.Selenium.Keys.Enter);
                    action.DoubleClick(driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[1]/div/div[5]/div/form/div/div/div[3]/div/div[1]/table/tbody/tr/td[4]/img[18]"))).Build().Perform();

                    // extra lines to receive focus
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("HaikuUploadImage")));

                    driver.FindElement(By.Id("HaikuUploadImage")).Clear();
                    // extra lines to receive focus
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");
                    driver.FindElement(By.Id("HaikuUploadImage")).SendKeys(latestScreenshotILC);

                    //// additional lines to ensure image path is supplied
                    //driver.FindElement(By.Id("HaikuUploadImage")).Clear();
                    //driver.FindElement(By.Id("HaikuUploadImage")).SendKeys("");

                    //IAlert alert = driver.SwitchTo().Alert();
                    //alert.SendKeys(latestScreenshotILC);
                    //alert.Accept();

                    Thread.Sleep(3000);

                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");

                    driver.FindElement(By.Id("e-dialog-insertimageprompt-insert")).Click();
                    Thread.Sleep(2000);
                    // ==============================================================================
                    //action.SendKeys("\n\nIT Security Status: \n\n").Perform();

                    //// Image attachment code

                    //string latestScreenshotITSec = RetrieveLatestScreenshot(@"C:\Screenshots\ITSec");

                    //// extra lines to receive focus
                    //driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //driver.SwitchTo().Frame("s_MainFrame");

                    //driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[1]/div/div[5]/div/form/div/div/div[3]/div/div[1]/table/tbody/tr/td[4]/img[18]")).Click();
                    ////action.SendKeys(OpenQA.Selenium.Keys.Tab);
                    ////action.SendKeys(OpenQA.Selenium.Keys.Enter);
                    //action.DoubleClick(driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[1]/div/div[5]/div/form/div/div/div[3]/div/div[1]/table/tbody/tr/td[4]/img[18]"))).Build().Perform();

                    //wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("HaikuUploadImage")));

                    //// extra lines to receive focus
                    //driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //driver.SwitchTo().Frame("s_MainFrame");

                    //driver.FindElement(By.Id("HaikuUploadImage")).Clear();
                    //Thread.Sleep(2000);
                    //// extra lines to receive focus
                    //driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //driver.SwitchTo().Frame("s_MainFrame");

                    //driver.FindElement(By.Id("HaikuUploadImage")).SendKeys(latestScreenshotITSec);
                    //Thread.Sleep(2000);

                    //// extra lines to receive focus
                    //driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //driver.SwitchTo().Frame("s_MainFrame");

                    //driver.FindElement(By.Id("e-dialog-insertimageprompt-insert")).Click();
                    //Thread.Sleep(2000);
                    // ==============================================================================

                    //driver.SwitchTo().DefaultContent();
                    //driver.SwitchTo().Frame("s_MainFrame");
                    //action.Click(driver.FindElement(By.Id("e-dialog-insertimageprompt-insert"))).Build().Perform();
                    //driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[2]/div[18]/table[1]/tbody/tr/td[1]")).Click();


                    // extra lines to receive focus

                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");

                    //Add Excel Sheet as attachment ==================================================
                    IWebElement attachmentElement = driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[2]/div[14]/table[1]/tbody/tr/td[6]"));
                    action.Click(attachmentElement).Build().Perform();

                    Thread.Sleep(5000);

                    IAlert alert = driver.SwitchTo().Alert();
                    alert.SendKeys(excelFilePath);
                    alert.Accept();
                    Thread.Sleep(3000);


                    // extra lines to receive focus

                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                    driver.SwitchTo().Frame("s_MainFrame");

                    try
                    {
                        driver.FindElement(By.Id("e-actions-mailedit-send-text")).Click();

                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("Modal dialog present"))
                        {
                            try
                            {   //Need to put this in try block as closing the browser results in a modal dialog that causes an exception
                                driver.Close();
                                driver.Dispose();

                            }
                            catch (Exception ex1)
                            {
                                LogError(ex1);
                                driver.Close();
                                driver.Dispose();
                                //throw;
                            }
                        }
                        //throw;
                    }


                    //driver.SwitchTo().Window(driver.WindowHandles[0]);
                    //driver.SwitchTo().Frame("s_MainFrame");
                    //action.SendKeys(@"C:\ILCDelinquency.xlsx").Build().Perform();
                    //action.SendKeys(OpenQA.Selenium.Keys.Enter).Build().Perform();



                    //IList<IWebElement> imgElements = driver.FindElements(By.TagName("img"));
                    //for (int i = 1; i < imgElements.Count; i++)
                    //{
                    //    //System.out.println("*********************************************");
                    //    //System.out.println(divElements.get(i).getText());
                    //    //Console.WriteLine(divElements[i].GetAttribute("title"));
                    //    if (imgElements[i].GetAttribute("id").Equals("e-actions-maileditwithdel-attach-icon"))
                    //    {

                    //        //spanElements[i].Click();
                    //        //driver.FindElement(By.Id("button-1175-btnEl")).Click();
                    //        action.DoubleClick(imgElements[i]).Build().Perform();
                    //        break;
                    //    }

                    //}

                    //driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/div/div/div[1]/div/div[1]/div/div[3]/div[2]/div[14]/table[1]/tbody/tr/td[6]")).Click();





                    // ===============================================================================

                    //IList<IWebElement> spanElements = driver.FindElements(By.TagName("span"));
                    //for (int i = 1; i < spanElements.Count; i++)
                    //{
                    //    //System.out.println("*********************************************");
                    //    //System.out.println(divElements.get(i).getText());
                    //    //Console.WriteLine(divElements[i].GetAttribute("title"));
                    //    if (spanElements[i].GetAttribute("innerHTML").Equals("Send"))
                    //    {
                    //        //spanElements[i].Click();
                    //        //driver.FindElement(By.Id("button-1175-btnEl")).Click();
                    //        action.DoubleClick(spanElements[i]).Build().Perform();
                    //        break;
                    //    }

                    //}

                }
                Thread.Sleep(10000);
                try
                {   //Need to put this in try block as closing the browser results in a modal dialog that causes an exception
                    driver.Close();
                    driver.Dispose();

                }
                catch (Exception ex)
                {
                    LogError(ex);
                    driver.Close();
                    driver.Dispose();
                    //throw;
                }

            }
            catch (Exception ex)
            {
                this.LogError(ex);
                //throw;
                if (iterationCount <= 3)
                {
                    iterationCount++;
                    RestartEmailSender();

                }
                else
                {
                    driver.Close();
                    driver.Dispose();
                }
            }


        }

        public void RestartEmailSender()
        {
            // Add code to close existing driver and resend email
            driver.Close();
            driver.Dispose();
            SendEmailINotes();
        }


        public IWebElement CaptureScreenshot()
        {
            try
            {
                ////Obtain ILC Dashboard link from config file

                Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                //string dashboardLink = config.AppSettings.Settings["dashboardLink"].Value;


                driver = new PhantomJSDriver();
                //driver = new InternetExplorerDriver();
                driver.Manage().Window.Maximize();
                driver.Navigate().GoToUrl(dashboardLink);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div")));
                string tempString = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div/table[2]/tbody/tr/th[2]")).GetAttribute("innerHTML");
                string excludedString = " &nbsp;";
                char[] newChar = excludedString.ToCharArray();
                string completedPer = tempString.Trim(newChar);
                completionPercentage = Convert.ToDouble(completedPer);

                if (completionPercentage == 100)
                {
                    if (ConfigurationManager.AppSettings["emailsent"] != "false")
                    {
                        Environment.Exit(0);
                    }
                    //Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                    config.AppSettings.Settings.Remove("emailsent");
                    //ConfigurationManager.AppSettings["emailsent"] = "true";
                    config.AppSettings.Settings.Add("emailsent", "true");
                    config.Save(ConfigurationSaveMode.Minimal);
                }

                IWebElement element = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div"));

                string fileName = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".jpg";
                Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;

                System.IO.Directory.CreateDirectory(screenshotPath);

                System.Drawing.Bitmap screenshot = new System.Drawing.Bitmap(new System.IO.MemoryStream(byteArray));
                System.Drawing.Rectangle croppedImage = new System.Drawing.Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
                screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat);
                screenshot.Save(String.Format(screenshotPath + @"\" + fileName, System.Drawing.Imaging.ImageFormat.Jpeg));

                //MessageBox.Show("Screenshot Captured!");

                driver.Close();
                driver.Dispose();

                return element;
            }
            catch (Exception ex)
            {
                //logger.Error(e.StackTrace + ' ' + e.Message);
                this.LogError(ex);

                return null;
            }
        }

        public IWebElement CaptureScreenshotITSec()
        {
            try
            {
                driver = new PhantomJSDriver();
                //driver = new InternetExplorerDriver();
                driver.Manage().Window.Maximize();
                driver.Navigate().GoToUrl("http://9.120.216.237/SMOTools/ITsecuritydashboard.php");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div")));
                string tempString = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div/table[2]/tbody/tr/th[2]")).GetAttribute("innerHTML");
                string excludedString = " &nbsp;";
                char[] newChar = excludedString.ToCharArray();
                string completedPer = tempString.Trim(newChar);
                completionPercentage = Convert.ToDouble(completedPer);

                if (completionPercentage == 100)
                {
                    if (ConfigurationManager.AppSettings["emailsentITSec"] != "false")
                    {
                        Environment.Exit(0);
                    }
                    Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                    config.AppSettings.Settings.Remove("emailsentITSec");
                    //ConfigurationManager.AppSettings["emailsentITSec"] = "true";
                    config.AppSettings.Settings.Add("emailsentITSec", "true");
                    config.Save(ConfigurationSaveMode.Minimal);
                }

                IWebElement element = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div[1]/div[2]/div[1]/div/div"));

                string fileName = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".jpg";
                Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;

                System.IO.Directory.CreateDirectory(@"C:\Screenshots\ITSec");

                System.Drawing.Bitmap screenshot = new System.Drawing.Bitmap(new System.IO.MemoryStream(byteArray));
                System.Drawing.Rectangle croppedImage = new System.Drawing.Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
                screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat);
                screenshot.Save(String.Format(@"C:\Screenshots\ITSec\" + fileName, System.Drawing.Imaging.ImageFormat.Jpeg));

                //MessageBox.Show("Screenshot Captured!");

                driver.Close();
                driver.Dispose();

                return element;
            }
            catch (Exception ex)
            {
                //logger.Error(e.StackTrace + ' ' + e.Message);
                this.LogError(ex);

                return null;
            }
        }

        public void PasteImage()
        {
            try
            {
                //driver.Navigate().GoToUrl(baseURL + "/verse?");
                //driver.FindElement(By.CssSelector("button.compose-button")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[3]/button[4]")).Click();
                // ERROR: Caught exception [ERROR: Unsupported command [selectFrame |  | ]]
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.


                //Check if Frame has loaded
                IList<IWebElement> iFrames = driver.FindElements(By.TagName("iframe"));


                if (iFrames.Count > 2)
                {


                    for (int i = 0; i < iFrames.Count; i++)
                    {
                        string frameTitle = iFrames[i].GetAttribute("title");
                        if (frameTitle.StartsWith("Select"))
                        {
                            driver.SwitchTo().Frame(i);
                            break;
                        }
                    }
                    //driver.SwitchTo().Frame(1);

                }
                else
                {
                    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[3]/button[4]")).Click();


                    for (int i = 0; i < iFrames.Count; i++)
                    {
                        string frameTitle = iFrames[i].GetAttribute("title");
                        if (frameTitle.StartsWith("Select"))
                        {
                            driver.SwitchTo().Frame(i);

                        }
                    }
                    //driver.SwitchTo().Frame(1);

                }
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/form/input")));

                string latestScreenshot = RetrieveLatestScreenshot(@"C:\Screenshots\ILC");

                driver.FindElement(By.XPath("/html/body/form/input")).Clear();
                //driver.FindElement(By.XPath("/html/body/form/input")).Click();

                driver.FindElement(By.XPath("/html/body/form/input")).SendKeys(latestScreenshot);
                // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]

                //switch to parent

                //driver.SwitchTo().ParentFrame();
                //driver.SwitchTo().DefaultContent();

                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver); // define Actions object

                //driver.FindElement(By.XPath("/html/body/form/input")).Click();

                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();

                action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                Thread.Sleep(5000);

                //driver.FindElement(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[1]/td/div[1]/table/tbody/tr[2]/td/a")).Click();

                //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[1]/td/div[3]/table/tbody/tr[1]/td/table/tbody/tr/td/div/div/div/input")));

                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();

                action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();


                //driver.FindElement(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a")).Click();
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                //throw;
                if (iterationCount <= 3)
                {
                    iterationCount++;
                    RestartEmailSender();
                }
                else
                {
                    driver.Close();
                    driver.Dispose();
                }
            }


        }

        public void PasteImageITSec()
        {
            try
            {
                //driver.Navigate().GoToUrl(baseURL + "/verse?");
                //driver.FindElement(By.CssSelector("button.compose-button")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[3]/button[4]")).Click();
                driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[3]/button[4]")).Click();    //Clicking twice due to known issue.

                // ERROR: Caught exception [ERROR: Unsupported command [selectFrame |  | ]]
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120)); //you can play with the time integer  to wait for longer.


                //Check if Frame has loaded
                IList<IWebElement> iFrames = driver.FindElements(By.TagName("iframe"));


                if (iFrames.Count > 2)
                {


                    for (int i = 0; i < iFrames.Count; i++)
                    {
                        string frameTitle = iFrames[i].GetAttribute("title");
                        if (frameTitle.StartsWith("Select"))
                        {
                            driver.SwitchTo().Frame(i);
                            break;
                        }
                    }
                    //driver.SwitchTo().Frame(1);

                }
                else
                {
                    driver.FindElement(By.XPath("/html/body/div[5]/div[3]/div[6]/div[1]/div[2]/div/div/div[2]/div[3]/div[2]/div[3]/button[4]")).Click();


                    for (int i = 0; i < iFrames.Count; i++)
                    {
                        string frameTitle = iFrames[i].GetAttribute("title");
                        if (frameTitle.StartsWith("Select"))
                        {
                            driver.SwitchTo().Frame(i);

                        }
                    }
                    //driver.SwitchTo().Frame(1);

                }
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/form/input")));

                string latestScreenshot = RetrieveLatestScreenshot(@"C:\Screenshots\ITSec");

                driver.FindElement(By.XPath("/html/body/form/input")).Clear();
                //driver.FindElement(By.XPath("/html/body/form/input")).Click();

                driver.FindElement(By.XPath("/html/body/form/input")).SendKeys(latestScreenshot);
                // ERROR: Caught exception [ERROR: Unsupported command [selectWindow | name=Verse_https://mail.notes.na.collabserv.com/verse | ]]

                //switch to parent

                //driver.SwitchTo().ParentFrame();
                //driver.SwitchTo().DefaultContent();

                OpenQA.Selenium.Interactions.Actions action = new OpenQA.Selenium.Interactions.Actions(driver); // define Actions object

                //driver.FindElement(By.XPath("/html/body/form/input")).Click();

                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();

                action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();

                Thread.Sleep(5000);

                //driver.FindElement(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[1]/td/div[1]/table/tbody/tr[2]/td/a")).Click();

                //wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[1]/td/div[3]/table/tbody/tr[1]/td/table/tbody/tr/td/div/div/div/input")));

                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();
                action.SendKeys(OpenQA.Selenium.Keys.Tab).Perform();

                action.SendKeys(OpenQA.Selenium.Keys.Enter).Perform();


                //driver.FindElement(By.XPath("/html/body/div[12]/table/tbody/tr/td/div/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a")).Click();
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                //throw;
                if (iterationCount <= 3)
                {
                    iterationCount++;
                    RestartEmailSender();
                }
                else
                {
                    driver.Close();
                    driver.Dispose();
                }
            }


        }

        public string RetrieveLatestScreenshot(string path)
        {
            try
            {
                var directory = new DirectoryInfo(path);
                var myFile = (from f in directory.GetFiles()
                              orderby f.LastWriteTime descending
                              select f).First();
                if (myFile != null)
                {
                    return myFile.FullName;

                }
                else
                {
                    LogErrorText("Custom Error: Unable to find last ILC screenshot");
                    return "";
                }
                //file.

            }
            catch (Exception ex)
            {
                this.LogError(ex);
                throw;
            }

        }

        public int GetIso8601WeekOfYear(DateTime time)
        {
            try
            {
                // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
                // be the same week# as whatever Thursday, Friday or Saturday are,
                // and we always get those right
                DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
                if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
                {
                    time = time.AddDays(3);
                }

                // Return the week of our adjusted day
                return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                throw;
            }

        }

        public void AddValue(string key, string value)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                config.AppSettings.Settings.Add(key, value);
                config.Save(ConfigurationSaveMode.Minimal);
            }
            catch (Exception ex)
            {
                LogError(ex);
                throw;
            }

        }

        public void TestEncryption()
        {
            PasswordSecurity newObj = new PasswordSecurity();
            //Encrption
            string encryptedString = EncryptString(ToSecureString("test"));

            //Decryption
            string password = DecryptString(encryptedString);


        }

        public bool AddPasswordInfo(string intranetID, string password)
        {
            try
            {

                if (ConfigurationManager.AppSettings[intranetID] != null)
                {
                    //DialogResult result = MessageBox.Show("ID information already exists. Password will be updated.", "Alert", MessageBoxButtons.OK);

                    //if (result == DialogResult.OK)
                    //{
                    Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                    config.AppSettings.Settings.Remove(intranetID);
                    //Encrypt password
                    string encryptedPassword = AesCryp.Encrypt(password);
                    config.AppSettings.Settings.Add(intranetID, encryptedPassword);
                    //config.Save(ConfigurationSaveMode.Minimal);
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    return true;
                    //}


                }

                else
                {
                    // Now do your magic..
                    Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

                    string encryptedPassword = EncryptString(ToSecureString(password));

                    config.AppSettings.Settings.Add(intranetID, encryptedPassword);
                    config.Save(ConfigurationSaveMode.Minimal);
                    return true;

                }

            }
            catch (Exception ex)
            {
                this.LogError(ex);
                return false;
                //throw;
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        #region PasswordSecurity ===========================================================

        static byte[] entropy = System.Text.Encoding.Unicode.GetBytes("Salt Is Not A Password");

        public static string EncryptString(System.Security.SecureString input)
        {
            byte[] encryptedData = System.Security.Cryptography.ProtectedData.Protect(
                System.Text.Encoding.Unicode.GetBytes(ToInsecureString(input)),
                entropy,
                System.Security.Cryptography.DataProtectionScope.CurrentUser);

            return Convert.ToBase64String(encryptedData);
        }

        public static string DecryptString(string encryptedData)
        {
            try
            {
                byte[] decryptedData = System.Security.Cryptography.ProtectedData.Unprotect(
                    Convert.FromBase64String(encryptedData),
                    entropy,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                //return ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData));
                return System.Text.Encoding.Unicode.GetString(decryptedData);
            }
            catch
            {
                //return new SecureString();
                return "";
            }
        }

        public static SecureString ToSecureString(string input)
        {
            SecureString secure = new SecureString();
            foreach (char c in input)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        public static string ToInsecureString(SecureString input)
        {
            string returnValue = string.Empty;
            IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
            try
            {
                returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr);
            }
            return returnValue;
        }
        #endregion PasswordSecurity

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            try
            {

                if (textBoxIntranetId.Text != "" && textBoxPassword.Text != "")
                {
                    string intranetID = textBoxIntranetId.Text;
                    string password = textBoxPassword.Text;
                    bool state = AddPasswordInfo(intranetID, password);

                    if (state)
                    {
                        MessageBox.Show("Password encrypted and saved successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogError(ex);
                throw;
            }

        }

        public bool CopyCredentialsFromWebConfig(string path, string username)
        {
            if (File.Exists(path))
            {
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = path;

                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                //string connStr_Con = config.ConnectionStrings.ConnectionStrings["AppDb"].ToString();
                //config.AppSettings[]
                //Console.WriteLine(connStr_Con);
                //Console.WriteLine();

                string encryptedPassword = config.AppSettings.Settings[username].Value;
                Console.WriteLine(encryptedPassword);
                Console.WriteLine();

                //string password = DecryptString(encryptedPassword);
                string password = AesCryp.Decrypt(encryptedPassword);

                addPasswordInfoSucc = AddPasswordInfo(username, password);

                if (addPasswordInfoSucc)
                {
                    return true;
                }

                else
                {
                    return false;
                }

            }
            else
            {
                return false;
            }
        }


        public void UpdateCredentialsFromWebConfig()
        {
            //emailSender = System.Configuration.ConfigurationManager.AppSettings["emailsender"];

            ExeConfigurationFileMap map = new ExeConfigurationFileMap();
            map.ExeConfigFilename = webConfigPath;

            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);

            foreach (string intranetId in managerIntranetIds)
            {
                if (config.AppSettings.Settings[intranetId] != null)
                {
                    bool isCopySucc = CopyCredentialsFromWebConfig(webConfigPath, intranetId);

                    // Can add code to handle failure of copy. But later.
                }
                else
                {
                    continue;
                }
            }


        }


        public void SetProperties()
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
                string managerIntranetIdsTemp = config.AppSettings.Settings["managerIntranetIDList"].Value;

                //Set managerIntranetIds string array
                managerIntranetIds = managerIntranetIdsTemp.Split(',');

                //Set webConfigPath string 
                webConfigPath = @config.AppSettings.Settings["webConfigPath"].Value;

                //New properties to set
                emailSender = config.AppSettings.Settings["emailSender"].Value;
                screenshotPath = config.AppSettings.Settings["screenshotPath"].Value;
                toList = config.AppSettings.Settings["toList"].Value;
                //added coded to remove &amp from to list if present. Added this LOC for Santosh span tolist
                toList = toList.Replace("&amp;", "&");

                excelFilePath = config.AppSettings.Settings["excelFilePath"].Value;
                errorFilePath = config.AppSettings.Settings["errorFilePath"].Value;
                dashboardLink = config.AppSettings.Settings["dashboardLink"].Value;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }


        }


        public void LogError(Exception ex)

        {

            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            message += string.Format("Message: {0}", ex.Message);

            message += Environment.NewLine;

            message += string.Format("StackTrace: {0}", ex.StackTrace);

            message += Environment.NewLine;

            message += string.Format("Source: {0}", ex.Source);

            message += Environment.NewLine;

            message += string.Format("TargetSite: {0}", ex.TargetSite.ToString());

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            string path = errorFilePath;

            FileInfo fi = new FileInfo(path);
            if (!fi.Exists)
            {
                File.Create(path).Dispose();    //with Dispose() it will give error that file is being used by other process.
            }

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }

        }

        public void LogErrorText(string messageText)

        {

            string message = string.Format("Time: {0}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            message += string.Format("Message: {0}", messageText);

            message += Environment.NewLine;

            message += Environment.NewLine;

            message += "-----------------------------------------------------------";

            message += Environment.NewLine;

            string path = errorFilePath;

            FileInfo fi = new FileInfo(path);
            if (!fi.Exists)
            {
                File.Create(path).Dispose();    //with Dispose() it will give error that file is being used by other process.
            }

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(message);
                writer.Close();
            }

        }

        private void CloseAllBrowserWindows()
        {
            Process[] AllProcesses = Process.GetProcesses();
            foreach (var process in AllProcesses)
            {
                if (process.MainWindowTitle != "")
                {
                    string s = process.ProcessName.ToLower();
                    if (s == "iexplore")
                        process.Kill();
                }
            }
        }
    }
}
