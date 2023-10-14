using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.IO.Packaging;
using System.Windows.Forms;

namespace BlueIQ_Neuware
{
    public partial class Form1 : Form
    {
        private IWebDriver? driver;
        private WebDriverWait? wait;
        private ExcelPackage? package;

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void SetupWebDriver()
        {
            var service = FirefoxDriverService.CreateDefaultService();
            UpdateUI(() => service.HideCommandPromptWindow = true);

            driver = new FirefoxDriver(service);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 1
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                UpdateUI(() => excelFilePathTextBox.Text = openFileDialog.FileName);
            }
        }

        private async void StartButton_Click(object sender, EventArgs e)
        {
            string username = usernameTextBox.Text;
            string password = passwordTextBox.Text;
            string excelFilePath = excelFilePathTextBox.Text;
            string location = locationTextBox.Text;

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password) || string.IsNullOrWhiteSpace(excelFilePath) || string.IsNullOrWhiteSpace(location))
            {
                // Display a message box to inform the user, specifying the parent form
                var messageBoxResult = MessageBox.Show(this, "Please fill in all required fields (Username, Password, and Excel File).", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                return; // Exit the method early
            }
            // Update the status label to show "Logging in"
            UpdateUI(() => statusLabel.Text = "Logging in");

            bool loginSuccess = await Task.Run(() => LoginToSite(username, password));
            try
            {
                if (loginSuccess)
                {
                    //ShowMessage("Login successful");
                    UpdateUI(() => statusLabel.Text = "Logged in");

                    // Process the Excel file after successful login
                    await Task.Run(() => ProcessExcelFile(excelFilePath, location));
                    UpdateUI(() => statusLabel.Text = "Devices have been added");
                    ShowMessage("The devices have been booked, please check the excel file to confirm");
                }
                else
                {
                    ShowMessage("Login failed");
                    UpdateUI(() => statusLabel.Text = "Login failed");
                }
            }
            catch (Exception ex)
            {
                // Handle the exception
                ShowMessage("An error occurred: " + ex.Message);
                UpdateUI(() => statusLabel.Text = "Error occurred");
            }
            finally
            {
                usernameTextBox.Text = "";
                passwordTextBox.Text = "";
                excelFilePathTextBox.Text = "";
                locationTextBox.Text = "";
            }
        }

        public bool LoadWebsite()
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

        public bool WaitForElementToDisappear(string elementXpath, int retries = 3)
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
                        UpdateUI(() => statusLabel.Text = "Retrying..      .");
                }
            }
            UpdateUI(() => statusLabel.Text = "Page took too long to load");
            return false;
        }

        public bool LoadPage(string page)
        {
            driver.Navigate().GoToUrl(page);
            return LoadWebsite();
        }

        public bool LoginToSite(string username, string password)
        {
            if (!LoadPage(BlueDictionary.LINKS["LOGIN"]))
                return false;

            try
            {
                IWebElement usernameElem = wait.Until(driver => driver.FindElement(By.XPath(BlueDictionary.LOGIN_PAGE["USERNAME"])));
                IWebElement passwordElem = wait.Until(driver => driver.FindElement(By.XPath(BlueDictionary.LOGIN_PAGE["PASSWORD"])));
                usernameElem.SendKeys(username);
                passwordElem.SendKeys(password);

                IWebElement loginButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.LOGIN_PAGE["BUTTON"])));
                loginButton.Click();

                if (!WaitForElementToDisappear(BlueDictionary.LOGIN_PAGE["LOADING"]))
                    return false;
                UpdateUI(() => statusLabel.Text = "Loading done");
                try
                {
                    // Create a separate WebDriverWait instance with a shorter timeout
                    var customWait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

                    IWebElement errorMessageElem = customWait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath(BlueDictionary.LOGIN_PAGE["LOGIN_ERROR"])))[0];
                    string errorMessage = errorMessageElem.Text.Trim();
                    UpdateUI(() => statusLabel.Text = "Error read");

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
                Console.Error.WriteLine(e.ToString());
                return false;
            }
        }

        public bool AddPallet(string pallet)
        {
            try
            {
                UpdateUI(() => statusLabel.Text = "Adding Pallet");
                driver.Navigate().GoToUrl(BlueDictionary.LINKS["AUDIT"]); // Replace with your actual URL
                var palletElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["PALLET_ID"]))); // Replace with your actual XPath
                palletElement.SendKeys(pallet);
                var checkElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_PALLET"]))); // Replace with your actual XPath
                checkElement.Click();
                return true;
            }
            catch (Exception)
            {
                // Maybe log the error or update a label on your form
                // statusLabel.Text = e.Message;
                return false;
            }
        }

        private void ProcessExcelFile(string excelFilePath, string location)
        {
            int counter = 0;
            string oldPallet = "";
            bool newPallet = false;
            Dictionary<string, object> data = new();

            package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelFilePath));
            var ws = package.Workbook.Worksheets[0]; // Access package from the class level
            int maxColumn = ws.Dimension.End.Column;

            // Find the last row with data
            int rowCount = ws.Cells[ws.Dimension.Address].Rows;

            UpdateUI(() => progressBar.Maximum = rowCount - 2); // Deducting 2 as you're starting from the second row and excluding the header row.
            UpdateUI(() => progressBar.Value = 0); // Reset the progress bar at the start.

            for (int row = 2; row <= rowCount; row++)
            {
                // Check if the row is empty (assuming column 2 is the part number column)
                if (string.IsNullOrEmpty(ws.Cells[row, 2].Text))
                {
                    // Skip empty rows
                    continue;
                }

                UpdateUI(() => statusLabel.Text = "Booking next device");
                data["part_number"] = ws.Cells[row, 2].Text;
                data["serial"] = ws.Cells[row, 3].Text;
                data["pallet"] = ws.Cells[row, 4].Text;
                data["pono"] = ws.Cells[row, 5].Text;

                if (counter < 2)
                {
                    // For debugging purposes
                    ShowMessage($"Part Number: {data["part_number"]}, Serial: {data["serial"]}, Pallet: {data["pallet"]}, PONO: {data["pono"]}");
                    DialogResult result = MessageBox.Show("Want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    if (result == DialogResult.No)
                    {
                        return; // Exit the function if the user chooses "No"
                    }
                }

                if (oldPallet != data["pallet"].ToString())
                {
                    newPallet = true;
                }

                try
                {
                    if (newPallet && !string.IsNullOrEmpty(data["pallet"].ToString()))
                    {
#pragma warning disable CS8604 // Possible null reference argument.
                        bool isPalletAdded = AddPallet(data["pallet"].ToString());
#pragma warning restore CS8604 // Possible null reference argument.
                        if (!isPalletAdded)
                        {
                            oldPallet = "";
                            ws.Cells[row, maxColumn + 1].Value = "Pallet Error";
                        }
                    }
                    bool isDeviceAdded = AddDevice(data, newPallet, location); // Assuming `driver` and `wait` are accessible here
                    ws.Cells[row, maxColumn + 1].Value = isDeviceAdded ? "Done" : "Error";
                    package.Save(); // Save changes to the Excel file after every row update
                    newPallet = false;
                    if (!string.IsNullOrEmpty(data["pallet"].ToString()))
                    {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                        oldPallet = data["pallet"].ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                    }
                }
                catch (Exception)
                {
                    continue;
                }
                finally
                {
                    if (counter < 2)
                    {
                        ShowMessage($"Data Processed: {string.Join(", ", data.Values)}");
                        counter++;
                    }
                    UpdateUI(() =>
                    {
                        progressBar.Value += 1;

                        // Calculate and update the percentage label
                        int percentage = (int)(((double)progressBar.Value / (double)progressBar.Maximum) * 100);
                        percentLabel.Text = $"{percentage}%";
                    });
                }
            }
            package.Save();
        }



        private bool AddDevice(Dictionary<string, object> data, bool newPallet, string location)
        {
            try
            {
                if (!WaitForElementToDisappear(BlueDictionary.AUDIT_PAGE["LOADING"]))
                    return false;

                System.Threading.Thread.Sleep(800);

                var partEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"])));
                System.Threading.Thread.Sleep(800);

                partEl.SendKeys(data["part_number"].ToString());
                partEl.SendKeys(OpenQA.Selenium.Keys.Tab);

                if (!WaitForElementToDisappear(BlueDictionary.AUDIT_PAGE["LOADING"]))
                    return false;

                try
                {
                    var serialEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"])));
                    serialEl.SendKeys(data["serial"].ToString());
                }
                catch (Exception)
                {
                    partEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["PART_NUMBER"])));
                    partEl.SendKeys(data["part_number"].ToString());
                    partEl.SendKeys(OpenQA.Selenium.Keys.Tab);

                    if (!WaitForElementToDisappear(BlueDictionary.AUDIT_PAGE["LOADING"]))
                        return false;

                    var serialElRetry = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["SERIAL#"])));
                    serialElRetry.SendKeys(data["serial"].ToString());
                }

                var assetEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["ASSET"])));
                assetEl.SendKeys(BlueDictionary.ASSET.ToString());

                var weightEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["WEIGHT"])));
                weightEl.SendKeys(BlueDictionary.WEIGHT.ToString());

                if (newPallet)
                {
                    var locationEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["LOCATION"])));
                    locationEl.SendKeys(location);

                    var lockLocEl = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["LOCK_LOCATION"])));
                    lockLocEl.Click();
                }

                var warrantyEl = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["WARRANTY"])));
                warrantyEl.Click();

                var inBoxEl = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_IN_BOX"])));
                inBoxEl.Click();

                var newStockEl = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["NEW_STOCK"])));
                newStockEl.Click();

                var ponoEl = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(BlueDictionary.AUDIT_PAGE["PONO"])));
                ponoEl.SendKeys(data["pono"].ToString());

                System.Threading.Thread.Sleep(500);

                var saveEl = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(BlueDictionary.AUDIT_PAGE["SAVE"])));
                saveEl.Click();

                if (!WaitForElementToDisappear(BlueDictionary.AUDIT_PAGE["LOADING"]))
                    return false;

                var driverWindowHandles = driver.WindowHandles;
                driver.SwitchTo().Window(driverWindowHandles[1]);
                driver.Close();
                driver.SwitchTo().Window(driverWindowHandles[0]);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void UpdateUI(Action action)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(action);
            }
            else
            {
                action();
            }
        }

        public static void ShowMessage(string message)
        {
            MessageBox.Show(message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            // Bring the form to the front when the application starts
            this.BringToFront();
            SetupWebDriver();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // Check if the ExcelPackage is not null and the Excel file is open
            if (package != null && package.Workbook.Worksheets.Count > 0)
            {
                // Save the Excel file (if it has changes)
                package.Save();

                // Close the Excel package and release resources
                package.Dispose();
            }

            // Close the WebDriver if it exists
            driver?.Quit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (About aboutBox = new About())
            {
                aboutBox.ShowDialog(this);
            }
        }

        private void createExcelTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                var ws = excel.Workbook.Worksheets.Add("Neuware");

                // Headers
                ws.Cells["A1"].Value = "scan";
                ws.Cells["B1"].Value = "Part number";
                ws.Cells["C1"].Value = "Serial number";
                ws.Cells["D1"].Value = "Pallet";
                ws.Cells["E1"].Value = "PoNo";

                // Define the range for the table.
                var tableRange = ws.Cells["A1:E2"];

                // Create a table based on this range.
                var table = ws.Tables.Add(tableRange, "NeuwareTable");

                // Format the table with gray color.
                //table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium11;// Start with no style
                table.ShowHeader = true;
                table.ShowFilter = true;

                // If you want the entire table's rows (and not just headers) to be gray, use this:
                table.Range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                table.Range.Style.Fill.BackgroundColor.SetColor(Color.Gray);

                // Auto fit columns
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.FileName = "Neuware";
                saveFileDialog.DefaultExt = ".xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fi = new FileInfo(saveFileDialog.FileName);
                    excel.SaveAs(fi);
                }
            }
        }
    }
}