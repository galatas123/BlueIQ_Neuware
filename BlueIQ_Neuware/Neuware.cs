using OfficeOpenXml;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System.Threading;

namespace BlueIQ_Neuware
{
    internal class Neuware
    {
        // Define a delegate for the event
        public delegate void ProgressUpdateHandler(int value, string percentageText = "");
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void SetMaxProgressHandler(int maxValue);
        public delegate void MessageHandler(string message, MessageBoxIcon icon = MessageBoxIcon.Information);

        // Define the event using the delegate
        public static event ProgressUpdateHandler? ProgressUpdated;
        public static event StatusUpdateHandler? StatusUpdated;
        public static event SetMaxProgressHandler? SetMaxProgress;
        public static event MessageHandler? ShowMessage;

        public static void Start_neuware(CancellationToken cancellationToken)
        {
            // Inside the method, you can periodically check if cancellation has been requested
            
            // Ensure the driver and wait objects from global_functions are initialized
            if (Global_functions.driver == null || Global_functions.wait == null)
            {
                throw new InvalidOperationException("WebDriver or WebDriverWait not initialized.");
            }

            StartBooking(cancellationToken);
        }

        private static void StartBooking(CancellationToken cancellationToken)
        {
            int progressBarValue = 0;
            int progressBarMaximum = 0;
            int maxColumn = 2;
            bool newPallet = true;
            Dictionary<string, object> data = new();
            var ws = Global_functions.package.Workbook.Worksheets[0]; // Access package from the class level
            

            // Find the last row with data
            int rowCount = ws.Cells[ws.Dimension.Address].Rows - 1;
            Global_functions.CreateJob(rowCount, true);
            if (JobInfo.Current.PalletId == "")
                throw new Exception("Failed to create job");
            SetMaxProgress?.Invoke(rowCount); // Deducting 2 as you're starting from the second row and excluding the header row.
            progressBarMaximum = rowCount;
            ProgressUpdated?.Invoke(0); // Reset the progress bar at the start.

            for (int row = 2; row <= rowCount + 1; row++)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }

                StatusUpdated?.Invoke(Languages.Resources.BOOK_NEXT);
                data["part_number"] = ws.Cells[row, 2].Text;
                data["serial"] = ws.Cells[row, 1].Text;

                try
                {
                    
                    if (newPallet)
                    {
                        if (!Global_functions.AddPallet())
                        {
                            ws.Cells[row, maxColumn].Value = "Pallet Error";
                            continue;
                        }
                    }
                    if ((data["serial"].ToString().Length == 8) && (data["part_number"].ToString().Length == 7))
                    {
                        if (!AddDevice(data, newPallet, ws, row, maxColumn))
                        {
                            Global_functions.driver.Navigate().Refresh();
                            newPallet = true;
                            continue;
                        }
                    }
                    else
                    {
                        if (data["serial"].ToString().Length == 8)
                            ws.Cells[row, maxColumn].Value = "Serial not 8 digits long";
                        else if (data["separt_numberrial"].ToString().Length == 7)
                            ws.Cells[row, maxColumn].Value = "part number not 7 digits long";
                        continue;
                    }
                    newPallet = false;
                }
                catch (Exception ex)
                {
                    Global_functions.LogError(Global_functions.GetCallerFunctionName(), (ex.ToString()));
                    ws.Cells[row, maxColumn].Value = "error check logs";
                    continue;
                }
                finally
                {
                    if (progressBarValue < rowCount + 1)
                    {
                        progressBarValue++;
                        int percentage = (int)(((double)progressBarValue / (double)progressBarMaximum) * 100);
                        ProgressUpdated?.Invoke(progressBarValue, $"{percentage}%");
                    }
                    Global_functions.package.Save();
                }
            }
            Global_functions.package.Save();
        }

        private static bool AddDevice(Dictionary<string, object> data, bool newPallet, ExcelWorksheet ws, int row, int maxColumn)
        {
            try
            {
                bool isElementClickable = false;
                int maxRetries = 3;
                int currentRetry = 0;

                Global_functions.WaitForLoadingToDisappear();
                while (!isElementClickable && currentRetry < maxRetries)
                {
                    try
                    {
                        Global_functions.wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"])));
                        isElementClickable = true; // Exit the loop once the element is clickable
                    }
                    catch (WebDriverTimeoutException)
                    {
                        currentRetry++; // Increase the retry counter
                        Global_functions.driver.Navigate().Refresh();
                    }
                }

                if (!isElementClickable)
                {
                    throw new Exception("save button did not appear on audit page");
                }

                Global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), data["part_number"].ToString());
                Global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), OpenQA.Selenium.Keys.Tab);

                var partNumber_el = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.AUDIT_PAGE["PART_TABLE"])));
                if (!partNumber_el.GetAttribute("value").Contains(data["part_number"]?.ToString() ?? ""))
                {
                    ws.Cells[row, maxColumn].Value = "check if part number is correct and if it exists in BlueIQ";
                    return false;
                }

                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"]), data["serial"]?.ToString(), true);
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["ASSET"]), BlueDictionary.ASSET.ToString(), true);
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["WEIGHT"]), BlueDictionary.WEIGHT.ToString(), true);

                if (newPallet)
                {
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCATION"]), JobInfo.Current.Location);
                    Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_LOCATION"]));
                }

                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["WARRANTY"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_IN_BOX"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_STOCK"]));
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["PONO"]), JobInfo.Current.JobOrPoNO);

                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"]), false);

                string alertMessage = Global_functions.HandleAlert();
                if (alertMessage != "")
                {
                    ws.Cells[row, maxColumn].Value = alertMessage;
                    return false;
                }

                if (!Global_functions.TryCloseSecondTab())
                {
                    Global_functions.LogError(Global_functions.GetCallerFunctionName(), "Failed to close the second tab after multiple attempts.");
                    ws.Cells[row, maxColumn].Value = "Failed to close second tab";
                    return false;
                }

                ws.Cells[row, maxColumn].Value = "Done";
                return true;
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), (ex.ToString()));
                ws.Cells[row, maxColumn].Value = "error check logs";
                return false;
            }
        }
    }
}