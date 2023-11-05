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
        public static ExcelPackage? package;

        public delegate void StatusUpdateHandler(string statusMessage);
        public static event StatusUpdateHandler? StatusUpdated;


        public static bool LoginToSite(string excelFilePath, string username, string password)
        { 
            SetupWebDriver();
            var customWait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
                    try
                    {
                        customWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.LOGIN_PAGE["LOCATION_POPUP"]))).Click();
                        return true;
                    }
                    catch(WebDriverTimeoutException)
                    {
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                LogError(GetCallerFunctionName(), (e.ToString()));
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
            waitAlert = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
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
            try
            {
                // Wait for the alert to be present.
                IAlert alert = waitAlert.Until(ExpectedConditions.AlertIsPresent());
                // Retrieve the alert's message.
                string alertMessage = alert.Text;
                LogError(GetCallerFunctionName(), alertMessage);
                // Accept the alert.
                alert.Accept();
                WaitForLoadingToDisappear();
                return alertMessage;
            }
            catch (WebDriverTimeoutException)
            {
                return "";
            }
            catch (NoAlertPresentException)
            {
                return "";
            }
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


        public static void SelectDropdownByVisibleText(By by, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new Exception(GetCallerFunctionName() + "Text provided is null or whitespace.");
            }

            try
            {
                System.Threading.Thread.Sleep(250);
                var selectElement = new SelectElement(wait.Until(ExpectedConditions.ElementToBeClickable(by)));
                selectElement.SelectByText(text);
                WaitForLoadingToDisappear();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void CheckAllInGridView()
        {
            try
            {
                // Wait for the JavaScript and AJAX calls to complete
                new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(
                    d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete")
                );

                // Execute the JavaScript function that checks all checkboxes in the DevExpress GridView
                ((IJavaScriptExecutor)driver).ExecuteScript(
                    "window['ctl00_ctl00_MainContent_PageMainContent_gvQuarantineAssets_header0_SelectAllCheckBox'].SetChecked(true); " +
                    "gvQuarantineAssets.SelectAllRowsOnPage(true);"
                );
            }
            catch (Exception)
            {
                throw;
            }
        }


        public static void CreateJob(int totaldevices, bool load = false)
        {
            try
            {
                StatusUpdated.Invoke("Receiving Job");
                const string AUDIO = "Audio";
                const string MISC = "Miscellaneous";
                const string SORT = "SORT";
                string totalweight = (totaldevices + 30).ToString() ;

                DateTime today = DateTime.Now;
                string formattedDate = today.ToString("dd/MM/yyyy");
                string formattedTime = today.ToString("hh:mm tt", new System.Globalization.CultureInfo("en-US"));
                driver.Navigate().GoToUrl(BlueDictionary.LINKS["RECEIVING"]);
                
                if (load)
                {
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["PONO"]), JobInfo.Current.JobOrPoNO);
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["PONO"]), OpenQA.Selenium.Keys.Tab);
                }
                else
                {
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["JOB_ID"]), JobInfo.Current.JobOrPoNO);
                    SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["JOB_ID"]), OpenQA.Selenium.Keys.Tab);
                }
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["OTHER"]));

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["BOL"]), "NA");
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["QTY_PALLET"]), "1");
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["DATE"]), formattedDate);
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["ARRIVAL_TIME"]), formattedTime);

                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["SAVE&EXIT"]), false);
                System.Threading.Thread.Sleep(500);
                HandleAlert();

                JobInfo.Current.PalletId = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.RECEIVING_PAGE["PALLET_ID"]))).Text;
                JobInfo.Current.LoadId = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.RECEIVING_PAGE["LOAD_ID"]))).Text;
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["PENCIL"]));

                driver.SwitchTo().Frame("ctl00_ctl00_MainContent_PageMainContent_ASPxPopupControlEditPallet_CIF-1");

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["WEIGHT"]), totalweight);
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["DUNNAGE_PALLET"]));

                SelectDropdownByVisibleText(By.Id("ddlClass"), AUDIO);
                SelectDropdownByVisibleText(By.Id("ddlCategory"), MISC);
                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["QTY_DEVICES"]), totaldevices.ToString());
                SelectDropdownByVisibleText(By.Id("ddlMoveToSite"), SORT);

                SendKeysToElement(By.XPath(BlueDictionary.RECEIVING_PAGE["LOCATION"]), JobInfo.Current.Location, true);
                ClickElement(By.XPath(BlueDictionary.RECEIVING_PAGE["SAVE"]));

                driver.SwitchTo().DefaultContent();
            }
            catch (Exception ex)
            {
                // Here you can log the exception or take other appropriate actions.
                LogError(GetCallerFunctionName(), (ex.ToString()));
            }
        }

        public static bool AddPallet()
        {
            try
            {
                StatusUpdated?.Invoke("Adding Pallet");
                if (!LoadPage(BlueDictionary.LINKS["AUDIT"])) // Replace with your actual URL
                    return false;

                string pallet_id = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["PALLET_ID"]))).Text;
                if (string.IsNullOrEmpty(pallet_id))
                {
                    SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["PALLET_ID"]), JobInfo.Current.PalletId);
                    ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_PALLET"]));
                }
                return true;
            }
            catch (Exception ex)
            {
                LogError(GetCallerFunctionName(), (ex.ToString()));
                return false;
            }
        }

        public static bool RemoveFromQuarantine()
        {
            try
            {
                if (!LoadPage(BlueDictionary.LINKS["QUARANTINE"]))
                {
                    return false;
                }

                SendKeysToVisibleElement(By.Id(BlueDictionary.QUARANTINE_PAGE["JOB#"]), JobInfo.Current.JobOrPoNO);
                Thread.Sleep(2000);
                ClickElement(By.Id(BlueDictionary.QUARANTINE_PAGE["SEARCH_BTN"]), false);
                Thread.Sleep(2000);
                string alertText = "";
                while (!alertText.Contains("Select at least one asset."))
                {
                    CheckAllInGridView();
                    ClickElement(By.Id(BlueDictionary.QUARANTINE_PAGE["RELEASE_BTN"]), false);
                    alertText = HandleAlert();
                    if (alertText.Contains("Select at least one asset."))
                    {
                        break;
                    }
                    WaitForLoadingToDisappear();
                    SendKeysToElement(By.Id(BlueDictionary.QUARANTINE_PAGE["RELEASE_REASON"]), JobInfo.Current.Mode);
                    ClickElement(By.Id(BlueDictionary.QUARANTINE_PAGE["RELEASE_YES"]), false);
                    alertText = HandleAlert(); // You need to define HandleAlert method
                }

                return true;
            }
            catch (Exception ex)
            {
                LogError(GetCallerFunctionName(), (ex.ToString()));
                throw;
            }
        }

        public static bool MassMove()
        {
            try
            {
                if (!LoadPage(BlueDictionary.LINKS["MASS_MOVE"]))
                {
                    return false;
                }

                SelectDropdownByVisibleText(By.Id(BlueDictionary.MASS_MOVE_PAGE["FROM_SITE_SEL"]), "SORT");
                SendKeysToElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["FROM_LOCATION"]), JobInfo.Current.Location);
                SelectDropdownByVisibleText(By.Id(BlueDictionary.MASS_MOVE_PAGE["TO_SITE_SEL"]), "FG(QUARANTINE)");
                SendKeysToElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["TO_LOCATION"]), BlueDictionary.LOCATIONS["QUARANTINE"]);
                ClickElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["MOVE_BTN"]));
                ClickElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["SUBMIT"]));
                ClickElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["QUAR_REASON_SEL"]));
                ClickElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["QUAR_REASON_OTHER"]));
                SendKeysToElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["QUAR_COMMENT"]), JobInfo.Current.Mode);
                ClickElement(By.Id(BlueDictionary.MASS_MOVE_PAGE["MOVE_BTN_POP"]));
                return true;
            }
            catch (Exception ex)
            {
                LogError(GetCallerFunctionName(), (ex.ToString()));
                throw;
            }
        }

        public static bool UpdateToProcessingCompleted()
        {
            try
            {
                if (!LoadPage(BlueDictionary.LINKS["LOAD"] + JobInfo.Current.LoadId))
                {
                    return false;
                }

                ClickElement(By.Id(BlueDictionary.LOAD_PAGE["LOAD_STATUS"]));
                IWebElement presortCompletedCheckbox = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id(BlueDictionary.LOAD_PAGE["PRESORT_COMPLETED"])));
                bool isPresortCompletedChecked = presortCompletedCheckbox.Selected;
                // Check if the 'OPS_COMPLETED' checkbox is checked
                IWebElement opsCompletedCheckbox = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id(BlueDictionary.LOAD_PAGE["OPS_COMPLETED"])));
                bool isOpsCompletedChecked = opsCompletedCheckbox.Selected;
                // If both checkboxes are checked, click on 'AUDIT_COMPLETED'
                if (isPresortCompletedChecked && isOpsCompletedChecked)
                {
                    ClickElement(By.Id(BlueDictionary.LOAD_PAGE["AUDIT_COMPLETED"]));
                }
                ClickElement(By.Id(BlueDictionary.LOAD_PAGE["SAVE"]), false);
                HandleAlert();

                if (!LoadPage(BlueDictionary.LINKS["JOBS"] + JobInfo.Current.JobOrPoNO))
                {
                    return false;
                }

                ClickElement(By.Id(BlueDictionary.JOBS_PAGE["SCHEDULES_TAB"]));
                SelectDropdownByVisibleText(By.Id(BlueDictionary.JOBS_PAGE["CARRIER"]), "Other");
                ClickElement(By.Id(BlueDictionary.JOBS_PAGE["JOB_INFO_TAB"]));
                SelectDropdownByVisibleText(By.Id(BlueDictionary.JOBS_PAGE["JOB_STATUS"]), "Processing Completed");
                ClickElement(By.Id(BlueDictionary.JOBS_PAGE["SAVE"]));
                SendKeysToElement(By.Id(BlueDictionary.JOBS_PAGE["REASON_NOTE"]), BlueDictionary.JOBS_PAGE["REASON"]);
                ClickElement(By.Id(BlueDictionary.JOBS_PAGE["REASON_OK"]), false);
                HandleAlert();
                return true;
            }
            catch (Exception ex)
            {
                LogError(GetCallerFunctionName(), (ex.ToString()));
                return false;
            }   
        }

        public static bool TryCloseSecondTab(int retries = 5)
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
                        LogError(GetCallerFunctionName(), (ex.ToString()));
                    }
                }

                // If second tab was not found or there was an issue, wait for a short duration before retrying.
                System.Threading.Thread.Sleep(500);
            }
            var driverWindowHandle = driver.WindowHandles;
            if (driverWindowHandle.Count > 1)
                return false;
            return true;
        }

        public static void LogError(string functionName, string errorMessage)
        {
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "errorLog.txt");
            using StreamWriter writer = new(logFilePath, true); // true means appending to the file
            string logEntry = $"\"{DateTime.Now:O}\", \"function: {functionName}\", \"Error: {errorMessage}\"";
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
                LogError(GetCallerFunctionName(), "Error reading settings file.");
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
                LogError(GetCallerFunctionName(), "Error updating settings file.");
                return;
            }

        }


        public static string GetCallerFunctionName([CallerMemberName] string memberName = "")
        {
            return memberName;
        }
    }
}