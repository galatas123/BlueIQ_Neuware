using OfficeOpenXml;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueIQ_Neuware
{
    internal class B2B
    {
        public delegate void ProgressUpdateHandler(int value, string percentageText = "");
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void SetMaxProgressHandler(int maxValue);
        public delegate void MessageHandler(string message);
        // Define the event using the delegate
        public static event ProgressUpdateHandler ProgressUpdated;
        public static event StatusUpdateHandler StatusUpdated;
        public static event SetMaxProgressHandler SetMaxProgress;
        public static event MessageHandler ShowMessage;

        public static void start_b2b(string location, string jobID, string creditType, CancellationToken cancellationToken)
        {
            // Inside the method, you can periodically check if cancellation has been requested
            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }
            // Ensure the driver and wait objects from global_functions are initialized
            if (global_functions.driver == null || global_functions.wait == null)
            {
                throw new InvalidOperationException("WebDriver or WebDriverWait not initialized.");
            }

            startBooking(location, jobID, creditType);
        }

        private static void startBooking(string location, string jobID, string creditType)
        {
            int counter = 0;
            int progressBarValue = 0;
            int progressBarMaximum = 0;
            bool newPallet = true;
            Dictionary<string, object> data = new();

            var ws = global_functions.package.Workbook.Worksheets[0]; // Access package from the class level
            int maxColumn = ws.Dimension.End.Column;

            // Find the last row with data
            int rowCount = ws.Cells[ws.Dimension.Address].Rows;
            string pallet_id = global_functions.createJob(jobID, rowCount - 1, location);
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

                StatusUpdated?.Invoke("Booking next device");
                data["part_number"] = ws.Cells[row, 2].Text;
                data["serial"] = ws.Cells[row, 1].Text;

                if (counter < 2)
                {
                    // For debugging purposes
                    ShowMessage?.Invoke($"Part Number: {data["part_number"]}, Serial: {data["serial"]}, Pallet: {pallet_id}, JobID: {jobID}");
                    DialogResult result = MessageBox.Show("Want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.No)
                    {
                        return; // Exit the function if the user chooses "No"
                    }
                }

                try
                {
                    if (newPallet)
                    {
                        if (!global_functions.AddPallet(pallet_id))
                        {
                            ws.Cells[row, maxColumn + 1].Value = "Pallet Error";
                            continue;
                        }
                    }
                    if (data["serial"].ToString().Length == 8)
                    {
                        if (!AddDevice(data, newPallet, location, creditType, ws, row, maxColumn))
                        {
                            ws.Cells[row, maxColumn + 1].Value = "add Device Error";
                        }
                    }
                    else
                    {
                        ws.Cells[row, maxColumn + 1].Value = "Serial not 8 digits long";
                        continue;
                    }
                    newPallet = false;
                }
                catch (Exception ex)
                {
                    global_functions.LogError(nameof(startBooking), (ex.ToString()));
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
                    global_functions.package.Save();
                }
            }
            global_functions.package.Save();
        }

        private static bool AddDevice(Dictionary<string, object> data, bool newPallet, string location, string creditType, ExcelWorksheet ws, int row, int maxColumn)
        {
            try
            {
                global_functions.WaitForLoadingToDisappear();
                System.Threading.Thread.Sleep(500);

                try
                {
                    global_functions.wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"])));
                }
                catch (WebDriverTimeoutException)
                {
                    global_functions.driver.Navigate().Refresh();
                }

                global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), data["part_number"].ToString());
                global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), OpenQA.Selenium.Keys.Tab);
                System.Threading.Thread.Sleep(500);

                try
                {
                    global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"]), data["serial"].ToString());
                }
                catch (Exception)
                {
                    global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), data["part_number"].ToString());
                    global_functions.SendKeysToVisibleElement(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"]), OpenQA.Selenium.Keys.Tab);
                    System.Threading.Thread.Sleep(500);
                    global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"]), data["serial"].ToString());
                }

                global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["ASSET"]), BlueDictionary.ASSET.ToString());
                global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["WEIGHT"]), BlueDictionary.WEIGHT.ToString());

                if (newPallet)
                {
                    global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCATION"]), location);
                    global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_LOCATION"]));
                }

                global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["WARRANTY"]));
                global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_IN_BOX"]));
                if (creditType == "full credit")
                {
                    global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERVICE_TYPE_CREDIT"]));
                    global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["APPROVED"]));
                }
                else
                {
                    global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERVICE_TYPE_SWAP"]));
                }
                
                System.Threading.Thread.Sleep(500);

                global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"]), false);

                try
                {

                    string alertMessage = global_functions.handleAlert();
                    // Update the Excel file with the alert message.
                    ws.Cells[row, maxColumn + 1].Value = alertMessage;

                    return false;
                }
                catch (WebDriverTimeoutException)
                {
                    global_functions.WaitForLoadingToDisappear();
                }

                global_functions.WaitForLoadingToDisappear();

                if (!global_functions.TryCloseSecondTab())
                {
                    global_functions.LogError(nameof(AddDevice), "Failed to close the second tab after multiple attempts.");
                    ws.Cells[row, maxColumn + 1].Value = "Failed to close second tab";
                    return false;
                }
                ws.Cells[row, maxColumn + 1].Value = "Done";
                return true;
            }
            catch (Exception ex)
            {
                global_functions.LogError(nameof(AddDevice), (ex.ToString()));
                ws.Cells[row, maxColumn + 1].Value = ex.ToString();
                return false;
            }
        }
    }
}
