using OfficeOpenXml;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;

namespace BlueIQ_Neuware
{
    internal class Neuware
    {
        // Define a delegate for the event
        public delegate void ProgressUpdateHandler(int value, string percentageText = "");
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void SetMaxProgressHandler(int maxValue);
        public delegate void MessageHandler(string message);

        // Define the event using the delegate
        public static event ProgressUpdateHandler? ProgressUpdated;
        public static event StatusUpdateHandler? StatusUpdated;
        public static event SetMaxProgressHandler? SetMaxProgress;
        public static event MessageHandler? ShowMessage;

        public static void Start_neuware(string location, string pono, CancellationToken cancellationToken)
        {
            // Inside the method, you can periodically check if cancellation has been requested
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }
            // Ensure the driver and wait objects from global_functions are initialized
            if (Global_functions.driver == null || Global_functions.wait == null)
            {
                throw new InvalidOperationException("WebDriver or WebDriverWait not initialized.");
            }

            StartBooking(location, pono);
        }

        private static void StartBooking(string location, string pono)
        {
            int counter = 0;
            int progressBarValue = 0;
            int progressBarMaximum = 0;
            int maxColumn = 2;
            bool newPallet = true;
            Dictionary<string, object> data = new();
            var ws = Global_functions.package.Workbook.Worksheets[0]; // Access package from the class level
            

            // Find the last row with data
            int rowCount = ws.Cells[ws.Dimension.Address].Rows;
            string pallet_id = Global_functions.CreateJob(pono, rowCount - 1, location, true);
            if (pallet_id == "")
                throw new Exception("Failed to create job");
            SetMaxProgress?.Invoke(rowCount - 1); // Deducting 2 as you're starting from the second row and excluding the header row.
            progressBarMaximum = rowCount - 1;
            ProgressUpdated?.Invoke(0); // Reset the progress bar at the start.

            for (int row = 2; row <= rowCount; row++)
            {
                // Check if the row is empty (assuming column 2 is the part number column)
                if (string.IsNullOrEmpty(ws.Cells[row, 2].Text))
                {
                    // Skip empty rows
                    continue;
                }

                StatusUpdated?.Invoke(Languages.Resources.BOOK_NEXT);
                data["part_number"] = ws.Cells[row, 2].Text;
                data["serial"] = ws.Cells[row, 1].Text;

                if (counter < 2)
                {
                    // For debugging purposes
                    ShowMessage?.Invoke($"Part Number: {data["part_number"]}, Serial: {data["serial"]}, Pallet: {pallet_id}, PONO: {pono}");
                    DialogResult? result = MessageBox.Show("Want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.No)
                    {
                        return; // Exit the function if the user chooses "No"
                    }
                }

                try
                {
                    
                    if (newPallet)
                    {
                        if (!Global_functions.AddPallet(pallet_id))
                        {
                            ws.Cells[row, maxColumn + 1].Value = "Pallet Error";
                            continue;
                        }
                    }
                    if ((data["serial"].ToString().Length == 8) && (data["part_number"].ToString().Length == 7))
                    {
                        AddDevice(data, newPallet, location, pono, ws, row, maxColumn);
                    }
                    else
                    {
                        if (data["serial"].ToString().Length == 8)
                            ws.Cells[row, maxColumn + 1].Value = "Serial not 8 digits long";
                        else if (data["separt_numberrial"].ToString().Length == 7)
                            ws.Cells[row, maxColumn + 1].Value = "part number not 7 digits long";
                        continue;
                    }
                    newPallet = false;
                }
                catch (Exception ex)
                {
                    Global_functions.LogError(Global_functions.GetCallerFunctionName(), (ex.ToString()));
                    ws.Cells[row, maxColumn + 1].Value = "error check logs";
                    continue;
                }
                finally
                {
                    if (counter < 2)
                    {
                        ShowMessage?.Invoke($"Data Processed: {string.Join(", ", data.Values)}");
                        counter++;
                    }
                    if (progressBarValue < 100)
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

        private static bool AddDevice(Dictionary<string, object> data, bool newPallet, string location, string pono, ExcelWorksheet ws, int row, int maxColumn)
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

                try
                {
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"]), data["serial"]?.ToString());
                }
                catch (Exception)
                {
                    Global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), data["part_number"].ToString());
                    Global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), OpenQA.Selenium.Keys.Tab);
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"]), data["serial"].ToString());
                }

                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["ASSET"]), BlueDictionary.ASSET.ToString());
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["WEIGHT"]), BlueDictionary.WEIGHT.ToString());

                if (newPallet)
                {
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCATION"]), location);
                    Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_LOCATION"]));
                }

                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["WARRANTY"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_IN_BOX"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_STOCK"]));
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["PONO"]), pono);

                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"]), false);

                try
                {
                    string alertMessage = Global_functions.HandleAlert();
                    // Update the Excel file with the alert message.
                    ws.Cells[row, maxColumn + 1].Value = alertMessage;

                    return false;
                }
                catch (WebDriverTimeoutException)
                {
                    Global_functions.WaitForLoadingToDisappear();
                }

                Global_functions.WaitForLoadingToDisappear();

                if (!Global_functions.TryCloseSecondTab())
                {
                    Global_functions.LogError(Global_functions.GetCallerFunctionName(), "Failed to close the second tab after multiple attempts.");
                    ws.Cells[row, maxColumn + 1].Value = "Failed to close second tab";
                    return false;
                }
                ws.Cells[row, maxColumn + 1].Value = "Done";
                return true;
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), (ex.ToString()));
                ws.Cells[row, maxColumn + 1].Value = "error check logs";
                return false;
            }
        }
    }
}