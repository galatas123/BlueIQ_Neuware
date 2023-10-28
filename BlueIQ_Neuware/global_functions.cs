using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Runtime.CompilerServices;

namespace BlueIQ_Neuware
{
    internal class Global_functions
    {
        public static IWebDriver? driver;
        public static WebDriverWait? wait;
        public static WebDriverWait? waitAlert;
        public static OfficeOpenXml.ExcelPackage? package;

        // Define a delegate for the event
        public delegate void StatusUpdateHandler(string statusMessage);

        // Define the event using the delegate
        public static event StatusUpdateHandler? StatusUpdated;

        public static bool LoginToSite(string excelFilePath, string username, string password)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            SetupWebDriver();
            if (!LoadPage(BlueDictionary.LINKS["LOGIN"]))
                return false;

            try
            {
                
                package = new ExcelPackage(new FileInfo(excelFilePath));
                SendKeysToVisibleElement(By.XPath(BlueDictionary.LOGIN_PAGE["USERNAME"]), username);
                SendKeysToVisibleElement(By.XPath(BlueDictionary.LOGIN_PAGE["PASSWORD"]), password);
                ClickElement(By.XPath(BlueDictionary.LOGIN_PAGE["BUTTON"]));

                if (!WaitForElementToDisappear(BlueDictionary.LOGIN_PAGE["LOADING"]))
                    return false;
                try
                {
                    // Create a separate WebDriverWait instance with a shorter timeout
                    var customWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

                    IWebElement errorMessageElem = customWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(BlueDictionary.LOGIN_PAGE["LOGIN_ERROR"])))[0];
                    string errorMessage = errorMessageElem.Text.Trim();

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
            StatusUpdated?.Invoke(Languages.Resources.LONG_LOAD);
            return false;
        }

        public static void WaitForLoadingToDisappear()
        {
            if (!WaitForElementToDisappear(BlueDictionary.loading))
                throw new Exception("Loading did not disappear after saving loading details.");
        }

        public static string HandleAlert()
        {
            IAlert alert = waitAlert.Until(ExpectedConditions.AlertIsPresent());

            // Retrieve the alert's message.
            string alertMessage = alert.Text;

            // Accept the alert.
            alert.Accept();
            return alertMessage;
        }

        public static void SendKeysToVisibleElement(By by, string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new Exception(GetCallerFunctionName() + "Text provided is null or whitespace.");
            }

            try
            {
                System.Threading.Thread.Sleep(250);
                var element = wait.Until(ExpectedConditions.ElementIsVisible(by));
                element.SendKeys(text);
                WaitForLoadingToDisappear();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void SendKeysToElement(By by, string? text, bool clearFirst = false)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new Exception(GetCallerFunctionName() + "Text provided is null or whitespace.");
            }

            try
            {
                System.Threading.Thread.Sleep(250);
                var element = wait.Until(ExpectedConditions.ElementToBeClickable(by));
                if (clearFirst)
                    element.Clear();
                element.SendKeys(text);
                WaitForLoadingToDisappear();
            }
            catch (Exception)
            {
                throw;
            }
        }


        public static void ClickElement(By by, bool waitForLoadAfterClick = true)
        {
            try
            {
                System.Threading.Thread.Sleep(250);
                var element = wait.Until(ExpectedConditions.ElementToBeClickable(by));
                element.Click();
                if (waitForLoadAfterClick)
                {
                    WaitForLoadingToDisappear();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string CreateJob(string referenceNumber, int totaldevices, string location, bool load = false)
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
                HandleAlert();

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
                return ""; // Return a default value or handle as appropriate.
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
            using StreamWriter writer = new(logFilePath, true); // true means appending to the file
            string logEntry = $"\"{DateTime.Now:O}\", \"{functionName}\", \"Error: {errorMessage}\"";
            writer.WriteLine(logEntry);
        }

        public static string GetSettings()
        {
            try
            {
                string settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");

                // Check if the file does not exist
                if (!File.Exists(settingsFilePath))
                {
                    // Create the file and write "language=" to it
                    File.WriteAllText(settingsFilePath, "language=");
                    return "en"; // return an empty string or default value if needed
                }
                else
                {
                    // If the file exists, read its content and extract the value after "language="
                    using StreamReader reader = new(settingsFilePath);
                    string? line = reader.ReadLine();
                    if (line != null && line.StartsWith("language="))
                    {
                        return line["language=".Length..];
                    }
                    else
                    {
                        return "en"; // return an empty string or default value if the expected content is not found
                    }
                }
            }
            catch (Exception)
            {
                LogError(nameof(GetSettings), "Error reading settings file.");
                return "en";
            }
           
        }

        public static void UpdateLanguageInSettingsFile(string language)
        {
            try
            {
                string settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");

                // Check if the file exists
                if (!File.Exists(settingsFilePath))
                {
                    // If the file doesn't exist, create it with the language parameter
                    File.WriteAllText(settingsFilePath, $"language={language}"); 
                    return;
                }

                // Read all lines from the file
                var lines = File.ReadAllLines(settingsFilePath).ToList();

                // Find the index of the line with the 'language=' prefix
                int index = lines.FindIndex(line => line.StartsWith("language="));

                // If found, update that line with the new language value
                if (index != -1)
                {
                    lines[index] = $"language={language}";
                }
                else
                {
                    // If not found, add the language setting to the end of the file
                    lines.Add($"language={language}");
                }

                // Write the modified lines back to the file
                File.WriteAllLines(settingsFilePath, lines);
            }
            catch (Exception)
            {
                LogError(nameof(UpdateLanguageInSettingsFile), "Error updating settings file."); 
                return;
            }
           
        }


        public static string GetCallerFunctionName([CallerMemberName] string memberName = "")
        {
            return memberName;
        }
    }
}