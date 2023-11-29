using ExcelDataReader;
using OfficeOpenXml;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System.Globalization;

namespace BlueIQ
{
    internal class Orders_report
    {
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void MessageHandler(string message, MessageBoxIcon icon = MessageBoxIcon.Information);
        public delegate void SaveFileHandler(out string? savedFilePath);
        public delegate void SetMaxProgressHandler(int maxValue);
        public delegate void ProgressUpdateHandler(int value, string percentageText = "");

        // Define the event using the delegate
        public static event StatusUpdateHandler? StatusUpdated;
        public static event MessageHandler? ShowMessage;
        public static event SaveFileHandler? RequestSaveFile;
        public static event SetMaxProgressHandler? SetMaxProgress;
        public static event ProgressUpdateHandler? ProgressUpdated;


        private static ExcelPackage? excelOutPackage;

        static readonly string InboundChannel = "eCollect/BoxIT";
        static readonly string OutboundChannel = "RETURN TO SOURCE";
        static readonly string InputExcel = "Orders.xls";
        static bool ignore = false;
        static string OrderType = "";
        static string ScanID = "";
        static string OutScanID = "";
        static string ShippedOut = "";
        static string Status = "";
        static string OrderID = "";
        static string Date = "";
        static string Comment = "";
        static string Serial = "";

        public static void StartReport(CancellationToken cancellationToken)
        {
            // Ensure the driver and wait objects from global_functions are initialized
            if (Global_functions.driver == null || Global_functions.wait == null)
            {
                throw new InvalidOperationException("WebDriver or WebDriverWait not initialized.");
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            CreateExcelOut("orders_übersicht.xlsx");
            if (DownloadExcel())
            {
                Itirate_list(cancellationToken);
            }

        }

        public static bool DownloadExcel()
        {
            try
            {
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["ORDER_FULLFILLMENT"]))
                    return false;
                Global_functions.ClickElement(By.Id(BlueDictionary.Q_ORDERS_PAGE["EXPORT_EXCEL"]));
                return true;
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
                ShowMessage?.Invoke(ex.Message, MessageBoxIcon.Error);
                return false;
            }
        }

        public static void Itirate_list(CancellationToken cancellationToken)
        {
            try
            {  
                var wsOut = excelOutPackage.Workbook.Worksheets[0]; 
                ProgressUpdated?.Invoke(0); // Reset the progress bar at the start.
                using (var stream = File.Open(InputExcel, FileMode.Open, FileAccess.Read))
                {
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    var result = reader.AsDataSet();
                    var dataTable = result.Tables[0];
                    int progressBarValue = 0;
                    int progressBarMaximum = 0;
                    SetMaxProgress?.Invoke(dataTable.Rows.Count);
                    progressBarMaximum = dataTable.Rows.Count;
                    for (int row = 1; row < dataTable.Rows.Count; row++) // Starting from 1 to skip header
                    //for (int row = 1; row < 2; row++)
                    {
                        try
                        {
                            if (cancellationToken.IsCancellationRequested)
                            {
                                return;
                            }

                            var dataRow = dataTable.Rows[row];
                            OrderID = dataRow[0].ToString();
                            Status = dataRow[7].ToString();
                            var channel = dataRow[2].ToString();

                            OrderType = channel.Equals(InboundChannel, StringComparison.OrdinalIgnoreCase) ? "INBOUND" : "OUTBOUND";
                            /*
                            OrderID = "1895643";
                            OrderType = "OUTBOUND";
                            Status = "SHIP STAGING";
                            */
                            if (Status.Equals("PICKING", StringComparison.OrdinalIgnoreCase))
                            {
                                ignore = true;
                                continue;
                            }

                            if (!GetOrderInfo(OrderID))
                            {
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Comment = "failed check logs";
                            Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
                        }
                        finally
                        {
                            if (progressBarValue < dataTable.Rows.Count + 1)
                            {
                                progressBarValue++;
                                int percentage = (int)(((double)progressBarValue / (double)progressBarMaximum) * 100);
                                ProgressUpdated?.Invoke(progressBarValue, $"{percentage}%");
                            }
                            if (!ignore)
                            {
                                var updateValues = new List<string> { OrderID, OrderType, Status, Date, ScanID, OutScanID, ShippedOut, Comment };
                                UpdateExcel(wsOut, updateValues);
                                excelOutPackage.Save();
                            }
                            ignore = false;
                            OrderID = OrderType = Status = ScanID = Date = Comment = OutScanID = ShippedOut = ""; 
                        }

                    }
                }
                excelOutPackage.Save();

                // Saving to user-selected path or deleting the original file
                string originalPath = excelOutPackage.File.FullName;
                string? userSelectedPath = null;
                RequestSaveFile?.Invoke(out userSelectedPath);

                try
                {
                    if (!string.IsNullOrEmpty(userSelectedPath))
                    {
                        excelOutPackage.SaveAs(new FileInfo(userSelectedPath));
                        File.Delete(originalPath);
                    }
                }
                catch (Exception ex)
                {
                    ShowMessage?.Invoke("Error deleting the original Excel file: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
                ShowMessage?.Invoke(ex.Message, MessageBoxIcon.Error);
            }
            finally
            {
                File.Delete(InputExcel);
                excelOutPackage?.Dispose();
            }
        }

        public static bool GetOrderInfo(string Order)
        {
            if (!Global_functions.LoadPage(BlueDictionary.LINKS["ORDER_DETAIL"] + Order))
                return false;

            Date = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["ORDER_DATE"]))).Text;
            if (DateTime.TryParseExact(Date, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                // Get the current date (without the time part)
                DateTime currentDate = DateTime.Today;

                // Calculate the difference in days
                TimeSpan dateDifference = currentDate - parsedDate;
                int daysOld = dateDifference.Days;
                Date = daysOld.ToString();
                if (daysOld < 2)
                {
                    ignore = true;
                }
            }

            if (OrderType == "INBOUND")
            {
                if (!GetScanID())
                {
                    return false;
                }         
            }

            if (OrderType == "OUTBOUND")
            {
                if (!GetOutboundInfo())
                {
                    return false;
                }
            }
            return true;
        }

        public static bool GetScanID()
        {
            try
            {
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["RECLAIM"]));
                Serial = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SERIAL_NO"]))).Text;
                if (!string.IsNullOrEmpty(Serial))
                {
                    if (!Global_functions.LoadPage(BlueDictionary.LINKS["AUDIT"]))
                        return false;
                    Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#_RADIO"]));
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), Serial);
                    Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), OpenQA.Selenium.Keys.Tab.ToString());
                    try
                    {
                        ScanID = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["SCAN_ID"]))).Text;
                    }
                    catch
                    {
                        Comment = "device not received";
                        return false;
                    }
                    return true;
                }
                else
                {
                    Comment = "no Inbound Serial";
                    return false;
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
                return false;
            }
        }

        public static bool GetOutboundInfo()
        {
            try
            {
                var info = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["SWAP_ORDER_REF"]))).GetAttribute("value");
                ScanID = info;
                if (info.Length > 0)
                {
                    string[] parts = info.Split('_');
                    if (parts.Length > 1)
                    {
                        ScanID = parts[1];
                    }
                }
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["TRACKING"]));
                ShippedOut = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SHIPPED_OUT"]))).Text;
                OutScanID = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["OUTBOUND_SCAN_ID"]))).Text;
                return true;
            }
            
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
                return false;
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
                    string[] headers = { "Order ID", "Type", "Status", "Age Days", "Inbound ScanID", "Outbound ScanID", "Shipped Outbound", "Comments" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        wsOut.Cells[1, i + 1].Value = headers[i];
                        wsOut.Cells[1, i + 1].Style.Font.Bold = true;
                        wsOut.Cells[1, i + 1].Style.Font.Size = 12;
                        wsOut.Cells[1, i + 1].Style.Font.Name = "Arial";
                        wsOut.Column(i + 1).Width = 20;
                    }
                    var tableRange = wsOut.Cells["A1:H400"];

                    // Create a table based on this range.
                    var table = wsOut.Tables.Add(tableRange, "Table");
                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Light9;
                    table.ShowFilter = true;
                    wsOut.Cells[wsOut.Dimension.Address].AutoFitColumns();

                    excelOutPackage.SaveAs(new FileInfo(excelToCreate));
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
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

