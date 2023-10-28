using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Text.RegularExpressions;


namespace BlueIQ_Neuware
{
    internal class Manual_outbound
    {
        // Define a delegate for the event
        public delegate void ProgressUpdateHandler(int value, string percentageText = "");
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void SetMaxProgressHandler(int maxValue);
        public delegate void MessageHandler(string message);
        public delegate void SaveFileHandler(out string? savedFilePath);

        // Define the event using the delegate
        public static event ProgressUpdateHandler? ProgressUpdated;
        public static event StatusUpdateHandler? StatusUpdated;
        public static event SetMaxProgressHandler? SetMaxProgress;
        public static event MessageHandler? ShowMessage;
        private static ExcelPackage? excelOutPackage;
        public static event SaveFileHandler? RequestSaveFile;

        public static void Start_manual_outbound(CancellationToken cancellationToken)
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
            CreateExcelOut("results.xlsx");
            Start_loop();
        }

        public static void Start_loop()
        {
            string oldScanId = "not defined";
            string oldOrderId = "not defined";
            string newScanId = "not defined";
            string newOrderId = "not defined";
            string newSerial = "not defined";
            int progressBarValue = 0;
            int progressBarMaximum = 0;

            var ws = Global_functions.package.Workbook.Worksheets[0];  // Utilize the already initialized excelPackage from global_functions
            var wsOut = excelOutPackage.Workbook.Worksheets[0]; // Assuming the package is initialized by createExcelOut
            int rowCount = ws.Cells[ws.Dimension.Address].Rows;
            SetMaxProgress?.Invoke(rowCount - 1); // Deducting 2 as you're starting from the second row and excluding the header row.
            progressBarMaximum = rowCount - 1;
            ProgressUpdated?.Invoke(0); // Reset the progress bar at the start.
            for (int row = 2; row <= rowCount; row++)
            {
                try
                {
                    StatusUpdated?.Invoke(Languages.Resources.NEXT_ORDER);
                    oldScanId = ws.Cells[row, 1].Text;
                    newSerial = ws.Cells[row, 2].Text;
                    Console.WriteLine($"new serial: {newSerial}, old scan: {oldScanId}");

                    newScanId = GetNewScanId(newSerial);
                    oldOrderId = GetInboundId(oldScanId);
                    newOrderId = CopyOrder(oldOrderId);
                    EditOutboundOrder(newOrderId, oldOrderId, oldScanId, newScanId);
                    FinalizeProcess(oldScanId, newScanId, newOrderId);

                    var updateValues = new List<string> { oldScanId, oldOrderId, newScanId, newOrderId, newSerial, "Success" };
                    UpdateExcel(wsOut, updateValues);
                    excelOutPackage.Save();  // Save the changes after updating the excel.

                    oldScanId = oldOrderId = newScanId = newOrderId = newSerial = "not defined";
                }
                catch (Exception ex)
                {
                    var errorValues = new List<string> { oldScanId, oldOrderId, newScanId, newOrderId, newSerial, "fail check logs" };
                    UpdateExcel(wsOut, errorValues);
                    Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
                    excelOutPackage.Save();  // Save the changes after updating the excel due to an error.

                    oldScanId = oldOrderId = newScanId = newOrderId = newSerial = "not defined";
                }
                finally
                {
                    if (progressBarValue < 100)
                    {
                        progressBarValue++;
                        int percentage = (int)(((double)progressBarValue / (double)progressBarMaximum) * 100);
                        ProgressUpdated?.Invoke(progressBarValue, $"{percentage}%");
                    }
                }
            }
            excelOutPackage.Save();
            // Store the original path of the excelOutPackage file
            string originalPath = excelOutPackage.File.FullName;
            string? userSelectedPath = null;
            RequestSaveFile?.Invoke(out userSelectedPath);

            if (!string.IsNullOrEmpty(userSelectedPath))
            {
                excelOutPackage.SaveAs(new FileInfo(userSelectedPath));
            }
            try
            {
                File.Delete(originalPath);
            }
            catch (Exception ex)
            {
                // Handle any issues encountered while trying to delete the file
                ShowMessage?.Invoke("Error deleting the original file: " + ex.Message);
            }

        }

        public static string GetNewScanId(string serial)
        {
            try
            {
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["AUDIT"]))
                    throw new Exception("Failed to load the audit page.");

                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#_RADIO"]));

                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), serial);
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), OpenQA.Selenium.Keys.Tab.ToString()); // Using .ToString() to convert the Keys.Tab to a string value.

                // check site and scan_id
                var siteElement = Global_functions.wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["SITE"])));
                var selectElement = new SelectElement(siteElement);
                var selectedOptionText = selectElement.SelectedOption.Text;

                if (selectedOptionText != "FG(CLIENT)")
                    throw new Exception($"Site is not FG(CLIENT) instead {selectedOptionText} is selected");

                var scanElement = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["SCAN_ID"])));
                if (string.IsNullOrEmpty(scanElement.Text))
                    throw new Exception("New scan ID is empty");

                return scanElement.Text;
            }
            catch (Exception ex)
            {
                var errorMsg = $"{Global_functions.GetCallerFunctionName()}: {ex.Message}";
                throw new Exception(errorMsg);
            }
        }

        public static string GetInboundId(string scanId)
        {
            try
            {
                // Navigate to old scan_id
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["REPAIR_DETAIL"] + scanId))
                    throw new Exception("Failed to load the page for scan ID: " + scanId);

                var notesText = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["REPAIR_NOTES"]))).Text;

                if (notesText.Contains("and added to Service Order"))
                {
                    var match = Regex.Match(notesText, @"Service Order[ {(]*(\d+)");
                    if (match.Success)
                        throw new Exception("An Outbound Order was already created for this scan id. Order ID: " + match.Groups[1].Value);
                }

                // Click return information
                Global_functions.ClickElement(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["RETURN_INFORMATION"]));

                // Switch to pop-up window
                var iframes = Global_functions.wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.TagName("iframe")));
                Global_functions.driver.SwitchTo().Frame(iframes[3]);

                var orderId = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["ORDER_ID"]))).Text;

                Global_functions.driver.SwitchTo().DefaultContent();

                return orderId;
            }
            catch (Exception ex)
            {
                var errorMsg = $"{Global_functions.GetCallerFunctionName()}: {ex.Message}";
                throw new Exception(errorMsg);
            }
        }

        public static string CopyOrder(string orderId)
        {
            try
            {
                // Navigate to order id
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["ORDER_DETAIL"] + orderId))
                    throw new Exception("Failed to load the page for order ID: " + orderId);

                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["COPY_ORDER"]));
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["CONFIRM1_YES"]));

                var newOrderIdElement = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["NEW_ORDER_ID"])));
                string newOrderId = newOrderIdElement.GetAttribute("value");

                Global_functions.WaitForLoadingToDisappear();
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["NAVIGATE_NO"]));

                return newOrderId;
            }
            catch (Exception ex)
            {
                var errorMsg = $"{Global_functions.GetCallerFunctionName()}: {ex.Message}";
                throw new Exception(errorMsg);
            }
        }

        public static void EditOutboundOrder(string newOrder, string oldOrder, string oldScan, string newScan)
        {
            try
            {
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["ORDER_DETAIL"] + newOrder))
                    throw new Exception($"Failed to load order detail page for order {newOrder}.");

                // Navigate and edit SHIP TO details
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SHIP_TO"]));
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["COMPANY_DROPDOWN"]));
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["FLENSBURG_OPTION"]));

                // Save SHIP TO details
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["SHIP_SAVE"]));

                // Frame operations
                Global_functions.driver.SwitchTo().Frame(BlueDictionary.ORDER_DETAILS_PAGE["FRAME_ADDRESS"]);
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["ADDRESS_CHECKBOX"]));
                Global_functions.ClickElement(By.Id("btnOK_CD"), false);
                Global_functions.HandleAlert();
                Global_functions.WaitForLoadingToDisappear();

                // Switching back to main content and refreshing
                try
                {
                    Global_functions.driver.SwitchTo().DefaultContent();
                    Global_functions.driver.Navigate().Refresh();
                }
                catch (Exception)
                {
                    Console.WriteLine("bad shit");
                }

                // Continue operations in the order details page
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SERVICE_TYPE"]));
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["SWAP"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["RETURN_TO_SOURCE"]));

                // Checkbox operations and submission
                var checkbox = Global_functions.wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["INCLUDE_LABEL_CHECKBOX"])));
                if (checkbox.Selected)
                    checkbox.Click();

                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SUBMIT_ORDER"]));
                Global_functions.ClickElement(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["CONFIRM_SUBMIT"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SAVE_ORDER"]), false);
                Global_functions.HandleAlert();
                Global_functions.WaitForLoadingToDisappear();

                // Swap reference and scan ID operations
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SWAP_ORDER_REF"]), $"{oldOrder}_{oldScan}");
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["ADD_SCAN"]));
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["INPUT_SCAN"]), $"{newScan}");
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["ADD"]));

                // More operations
                if (Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["MESSAGE"]))).Displayed)
                    System.Threading.Thread.Sleep(1000);

                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["CANCEL"]));
                var reclaim = Global_functions.wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["RECLAIM_LABEL"])));
                if (!reclaim.Selected)
                    reclaim.Click();
                System.Threading.Thread.Sleep(1000);
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["DELETE"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["CONFIRM_DEL"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SAVE_ORDER"]), false);
                Global_functions.HandleAlert();
                Global_functions.WaitForLoadingToDisappear();
            }
            catch (Exception ex)
            {
                var errorMsg = $"{Global_functions.GetCallerFunctionName()} : {ex.Message}";
                throw new Exception(errorMsg);
            }
        }

        public static bool FinalizeProcess(string oldScan, string newScan, string newOrder)
        {
            try
            {
                string note = $"Asset ({oldScan}) has been swapped with Asset ({newScan}) and added to Service Order ({newOrder})";
                Console.WriteLine(note);

                if (!Global_functions.LoadPage(BlueDictionary.LINKS["REPAIR_DETAIL"] + oldScan))
                    throw new Exception($"Failed to load the page for old scan: {oldScan}");

                // remove diag
                Global_functions.ClickElement(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["DIAG_TABLE"]));
                Global_functions.ClickElement(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["DIAG"]));

                // add note
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["REPAIR_INPUT"]), note);

                // save
                Global_functions.ClickElement(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["SAVE"]));
                System.Threading.Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                var errorMsg = $"{Global_functions.GetCallerFunctionName()} : {ex.Message}";
                throw new Exception(errorMsg);
            }
        }

        public static void CreateExcelOut(string excelToCreate)
        {
            try
            {
                if (File.Exists(excelToCreate))
                {
                    excelOutPackage = new ExcelPackage(new FileInfo(excelToCreate));
                }
                else
                {
                    excelOutPackage = new ExcelPackage(); // Create a new workbook
                    var wsOut = excelOutPackage.Workbook.Worksheets.Add("Sheet1"); // Add a new worksheet

                    // Add headers if creating a new workbook
                    string[] headers = { "Old_scanID", "Old_orderID", "New_scanID", "New_orderID", "New_Serial", "Result" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        wsOut.Cells[1, i + 1].Value = headers[i];
                        wsOut.Cells[1, i + 1].Style.Font.Bold = true;
                        wsOut.Cells[1, i + 1].Style.Font.Size = 12;
                        wsOut.Cells[1, i + 1].Style.Font.Name = "Arial";
                        wsOut.Column(i + 1).Width = 20;
                    }

                    excelOutPackage.SaveAs(new FileInfo(excelToCreate));
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
            }


        }

        public static IEnumerable<ExcelRangeBase[]> IterateRows(ExcelWorksheet ws, int minRow, int maxRow)
        {
            // Ensure maxRow is not greater than the actual number of rows in the worksheet
            int lastRow = ws.Dimension.End.Row;
            maxRow = Math.Min(maxRow, lastRow);

            for (int row = minRow; row <= maxRow; row++)
            {
                var range = ws.Cells[row, 1, row, ws.Dimension.End.Column];
                ExcelRangeBase[] cellsArray = new ExcelRangeBase[range.Columns];
                for (int col = 1; col <= range.Columns; col++)
                {
                    cellsArray[col - 1] = range[row, col];
                }
                yield return cellsArray;
            }
        }

        public static void UpdateExcel(ExcelWorksheet ws, List<string> values)
        {
            try
            {
                int nextRow = ws.Dimension.Rows + 1;  // next available row
                for (int idx = 0; idx < values.Count; idx++)
                {
                    var cell = ws.Cells[nextRow, idx + 1];
                    cell.Value = values[idx];
                    cell.Style.Font.Name = "Arial";
                    cell.Style.Font.Size = 11;
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
            }
        }
    }
}