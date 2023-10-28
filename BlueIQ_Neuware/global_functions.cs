using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;

namespace BlueIQ_Neuware
{
    internal class global_functions
    {
        public static IWebDriver? driver;
        public static WebDriverWait? wait;
        public static WebDriverWait? waitAlert;
        public static OfficeOpenXml.ExcelPackage package;
        // Define a delegate for the event
        public delegate void StatusUpdateHandler(string statusMessage);
        // Define the event using the delegate
        public static event StatusUpdateHandler StatusUpdated;

        public static bool LoginToSite(string username, string password, string excelFilePath)
        {
            SetupWebDriver();
            if (!LoadPage(BlueDictionary.LINKS["LOGIN"]))
                return false;

            try
            {
                package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelFilePath));
                IWebElement usernameElem = wait.Until(driver => driver.FindElement(By.XPath(BlueDictionary.LOGIN_PAGE["USERNAME"])));
                IWebElement passwordElem = wait.Until(driver => driver.FindElement(By.XPath(BlueDictionary.LOGIN_PAGE["PASSWORD"])));
                usernameElem.SendKeys(username);
                passwordElem.SendKeys(password);

                IWebElement loginButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.LOGIN_PAGE["BUTTON"])));
                loginButton.Click();

                if (!WaitForElementToDisappear(BlueDictionary.LOGIN_PAGE["LOADING"]))
                    return false;
                StatusUpdated?.Invoke("Loading done");
                try
                {
                    // Create a separate WebDriverWait instance with a shorter timeout
                    var customWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

                    IWebElement errorMessageElem = customWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(BlueDictionary.LOGIN_PAGE["LOGIN_ERROR"])))[0];
                    string errorMessage = errorMessageElem.Text.Trim();
                    StatusUpdated?.Invoke("Error read");

                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    return true;
                }

            }
            catch (Exception e)
            {
                LogError(nameof(LoginToSite), (e.ToString()));
                return false;
            }
        }

        public static void SetupWebDriver()
        {
            var service = FirefoxDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true; // Hide the command prompt window

            // Redirect logs to the system's "null" device
            service.FirefoxBinaryPath = "NUL";

            driver = new FirefoxDriver(service);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            waitAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
        }

        public static void CloseDriver()
        {
            package.Dispose();
            driver?.Quit();
        }

        public static bool LoadPage(string page)
        {
            driver.Navigate().GoToUrl(page);
            return LoadWebsite();
        }

        public static bool LoadWebsite()
        {
            try
            {
                wait.Until(d => (bool)((IJavaScriptExecutor)d).ExecuteScript("return document.readyState == 'complete'"));
                return true;
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        public static bool WaitForElementToDisappear(string elementXpath, int retries = 3)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath(elementXpath)));
                    return true;
                }
                catch (WebDriverTimeoutException)
                {
                    if (i < retries - 1)
                        StatusUpdated?.Invoke("Retrying..      .");
                }
            }
            StatusUpdated?.Invoke("Page took too long to load");
            return false;
        }

        public static void WaitForLoadingToDisappear()
        {
            if (!WaitForElementToDisappear(BlueDictionary.loading))
                throw new Exception("Loading did not disappear after saving loading details.");
        }

        public static string handleAlert()
        {
            IAlert alert = waitAlert.Until(ExpectedConditions.AlertIsPresent());

            // Retrieve the alert's message.
            string alertMessage = alert.Text;

            // Accept the alert.
            alert.Accept();
            return alertMessage;
        }

        public static void SendKeysToVisibleElement(By by, string text)
        {
            try
            {
                var element = wait.Until(ExpectedConditions.ElementIsVisible(by));
                element.SendKeys(text);
                WaitForLoadingToDisappear();
            }
            catch (Exception e)
            {
                LogError(GetCallerFunctionName(), e.Message);
                throw; // Re-throwing the exception to be handled or logged further up the call stack if needed
            }
        }

        public static void SendKeysToElement(By by, string text, bool clearFirst = false)
        {
            try
            {
                var element = wait.Until(ExpectedConditions.ElementToBeClickable(by));
                if (clearFirst)
                    element.Clear();
                element.SendKeys(text);
                WaitForLoadingToDisappear();
            }
            catch (Exception e)
            {
                LogError(GetCallerFunctionName(), e.Message);
                throw;
            }
        }

        public static void ClickElement(By by, bool waitForLoadAfterClick = true)
        {
            try
            {
                var element = wait.Until(ExpectedConditions.ElementToBeClickable(by));
                element.Click();
                if (waitForLoadAfterClick)
                {
                    WaitForLoadingToDisappear();
                }
            }
            catch (Exception e)
            {
                LogError(GetCallerFunctionName(), e.Message);
                throw;
            }
        }

        public static string createJob(string referenceNumber, int totaldevices, string location, bool load=false)
        {
            try
            {
                const string AUDIO = "Audio";
                const string MISC = "Miscellaneous";
                const string SORT = "SORT";
                string totalweight = totaldevices.ToString() + "30";

                DateTime today = DateTime.Now;
                string formattedDate = today.ToString("dd/MM/yyyy");
                string formattedTime = today.ToString("hh:mm tt");

                driver.Navigate().GoToUrl(BlueDictionary.LINKS["RECEIVING"]);
                if (load)
                {
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["LOAD_ID"]), referenceNumber);
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["LOAD_ID"]), OpenQA.Selenium.Keys.Tab);
                }
                else
                {
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["JOB_ID"]), referenceNumber);
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["JOB_ID"]), OpenQA.Selenium.Keys.Tab);
                }

                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["OTHER"]));

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["BOL"]), "NA");
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["QTY_PALLET"]), "1");
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["DATE"]), formattedDate);
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["ARRIVAL_TIME"]), formattedTime);

                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["SAVE&EXIT"]));
                handleAlert();

                var pallet = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.RECEIVING_PAGE["PALLET_ID"])));
                string pallet_id = pallet.Text;
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["PENCIL"]));

                driver.SwitchTo().Frame("ctl00_ctl00_MainContent_PageMainContent_ASPxPopupControlEditPallet_CIF-1");

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["WEIGHT"]), totalweight);
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["DUNNAGE_PALLET"]));

                new SelectElement(wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ddlClass")))).SelectByText(AUDIO);
                new SelectElement(wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ddlCategory")))).SelectByText(MISC);
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["QTY_DEVICES"]), totaldevices.ToString());
                new SelectElement(wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ddlMoveToSite")))).SelectByText(SORT);

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["LOCATION"]), location, true);
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["SAVE"]));

                driver.SwitchTo().DefaultContent();
                System.Environment.Exit(0);

                return pallet_id;
            }
            catch (Exception ex)
            {
                // Here you can log the exception or take other appropriate actions.
                LogError(nameof(AddPallet), (ex.ToString()));
                return null; // Return a default value or handle as appropriate.
            }
        }

        public static bool AddPallet(string pallet)
        {
            try
            {
                StatusUpdated.Invoke("Adding Pallet");
                driver.Navigate().GoToUrl(BlueDictionary.LINKS["AUDIT"]); // Replace with your actual URL
                SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["PALLET_ID"]), pallet);
                ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_PALLET"]));
                return true;
            }
            catch (Exception ex)
            {
                // Maybe log the error or update a label on your form
                LogError(nameof(AddPallet), (ex.ToString()));
                return false;
            }
        }

        public static bool TryCloseSecondTab(int retries = 3)
        {
            for (int i = 0; i < retries; i++)
            {
                var driverWindowHandles = driver.WindowHandles;
                if (driverWindowHandles.Count > 1)
                {
                    try
                    {
                        driver.SwitchTo().Window(driverWindowHandles[1]);
                        driver.Close();
                        driver.SwitchTo().Window(driverWindowHandles[0]);
                        return true; // Successfully closed the second tab.
                    }
                    catch (Exception ex)
                    {
                        LogError(nameof(TryCloseSecondTab), (ex.ToString()));
                    }
                }

                // If second tab was not found or there was an issue, wait for a short duration before retrying.
                System.Threading.Thread.Sleep(500);
            }
            // If we reach here, it means we've exhausted our retries and the second tab couldn't be closed.
            return false;
        }

        public static void LogError(string functionName, string errorMessage)
        {
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "errorLog.txt");
            using (StreamWriter writer = new StreamWriter(logFilePath, true)) // true means appending to the file
            {
                string logEntry = $"\"{DateTime.Now:O}\", \"{functionName}\", \"Error: {errorMessage}\"";
                writer.WriteLine(logEntry);
            }
        }

        private static string GetCallerFunctionName([CallerMemberName] string memberName = "")
        {
            return memberName;
        }

    }
}

