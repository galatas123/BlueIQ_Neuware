using OfficeOpenXml;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Range = Microsoft.Office.Interop.Excel.Range;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace BlueIQ_Neuware
{
    internal class orders_report
    {
        public delegate void StatusUpdateHandler(string statusMessage);
        public delegate void MessageHandler(string message, MessageBoxIcon icon = MessageBoxIcon.Information);
        public delegate void SaveFileHandler(out string? savedFilePath);

        // Define the event using the delegate
        public static event StatusUpdateHandler? StatusUpdated;
        public static event MessageHandler? ShowMessage;
        public static event SaveFileHandler? RequestSaveFile;


        private static ExcelPackage? package;
        private static ExcelPackage? excelOutPackage;

        static readonly string InboundChannel = "eCollect/BoxIT";
        static readonly string OutboundChannel = "RETURN TO SOURCE";
        static readonly string InputExcel = "Orders.xls";
        static string OrderType = "";
        static string ScanID = "";
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
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {   
                workbook = excelApp.Workbooks.Open(InputExcel);
                worksheet = (Worksheet)workbook.Worksheets[1]; // Assuming the first worksheet is the one you want
                Range range = worksheet.UsedRange;
                int rowCount = range.Rows.Count;
                var wsOut = excelOutPackage.Workbook.Worksheets[0]; // Assuming the package is initialized by createExcelOut
                //var ws = package.Workbook.Worksheets[0];
                //int rowCount = ws.Cells[ws.Dimension.Address].Rows - 1;
                for (int row = 2; row <= 5; row++)
                {
                    try
                    {
                        if (cancellationToken.IsCancellationRequested)
                        {
                            return;
                        }
                        /*
                        var order = ws.Cells[row, 1].Value.ToString();
                        var status = ws.Cells[row, 8].Value.ToString();
                        if (status == "PICKING")
                            continue;
                        if (ws.Cells[row, 3].Value.ToString() == InboundChannel)
                            OrderType = "INBOUND";
                        else
                            OrderType = "OUTBOUND";
                        if (!GetOrderInfo(OrderID))
                        {
                            continue;
                        }
                        */
                        Range orderCell = (Range)range.Cells[row, 1]; // Column 1 for 'order'
                        Range statusCell = (Range)range.Cells[row, 8]; // Column 8 for 'status'
                        Range channelCell = (Range)range.Cells[row, 3]; // Column 3 for 'channel'

                        OrderID = orderCell.Value?.ToString() ?? string.Empty;
                        Status = statusCell.Value?.ToString() ?? string.Empty;
                        string channel = channelCell.Value?.ToString() ?? string.Empty;
                        if (Status == "PICKING")
                            continue;
                        OrderType = channel == InboundChannel ? "INBOUND" : "OUTBOUND";
                        if (!GetOrderInfo(OrderID))
                        {
                            continue;
                        }
                        var updateValues = new List<string> { OrderID, OrderType, Status, ScanID, Date, Comment };
                        UpdateExcel(wsOut, updateValues);
                    }
                    catch (Exception ex)
                    {
                        var errorValues = new List<string> { OrderID, OrderType, Status, ScanID, Date, Comment, "fail check logs" };
                        UpdateExcel(wsOut, errorValues);
                        Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.Message);
                    }
                    finally
                    {
                        OrderID = OrderType = Status = ScanID = Date = Comment = "";
                        excelOutPackage.Save();
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
                excelOutPackage?.Dispose();
                package?.Dispose();
                workbook?.Close(false);
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        public static bool GetOrderInfo(string Order)
        {
            if (!Global_functions.LoadPage(BlueDictionary.LINKS["ORDER_DETAIL"] + Order))
                return false;
            Date = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.Id(BlueDictionary.ORDER_DETAILS_PAGE["ORDER_DATE"]))).Text;
            if (OrderType == "INBOUND")
            {
                Global_functions.ClickElement(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["RECLAIM"]));
                Serial = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.ORDER_DETAILS_PAGE["SERIAL"]))).Text;
                if (!GetScanID(Serial))
                    return false;
            }
            return true;
        }

        public static bool GetScanID(string serial)
        {
            try
            {
                if (!Global_functions.LoadPage(BlueDictionary.LINKS["AUDIT"]))
                    return false;
                Global_functions.ClickElement(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#_RADIO"]));
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), serial);
                Global_functions.SendKeysToElement(By.XPath(BlueDictionary.AUDIT_PAGE["SEARCH_ASSETS"]), OpenQA.Selenium.Keys.Tab.ToString());
                try
                {
                    ScanID = Global_functions.wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.REPAIR_DETAIL_PAGE["SCAN_ID"]))).Text;
                }
                catch
                {
                    Comment = "scan id not found";
                    return false;
                }
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
                    string[] headers = { "Order ID", "Type", "Status", "Inbound ScanID", "Creation Date", "Comments" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        wsOut.Cells[1, i + 1].Value = headers[i];
                        wsOut.Cells[1, i + 1].Style.Font.Bold = true;
                        wsOut.Cells[1, i + 1].Style.Font.Size = 12;
                        wsOut.Cells[1, i + 1].Style.Font.Name = "Arial";
                        wsOut.Column(i + 1).Width = 20;
                    }
                    var tableRange = wsOut.Cells["A1:F2"];

                    // Create a table based on this range.
                    var table = wsOut.Tables.Add(tableRange, "Table");
                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Light9;// Start with no style
                    table.ShowHeader = true;
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
