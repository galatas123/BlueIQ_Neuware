using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.IO.Packaging;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace BlueIQ_Neuware
{
    public partial class Form1 : Form
    {
        private string mode = "";
        private string creditType = "";
        private CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();

        public Form1()
        {
            InitializeComponent();
            global_functions.StatusUpdated += UpdateUI_statusLabel;
            Neuware.ProgressUpdated += UpdateUI_progressBar;
            Neuware.StatusUpdated += UpdateUI_statusLabel;
            Neuware.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            Neuware.ShowMessage += ShowMessage;
            B2B.ProgressUpdated += UpdateUI_progressBar;
            B2B.StatusUpdated += UpdateUI_statusLabel;
            B2B.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            B2B.ShowMessage += ShowMessage;
            manual_outbound.ProgressUpdated += UpdateUI_progressBar;
            manual_outbound.StatusUpdated += UpdateUI_statusLabel;
            manual_outbound.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            manual_outbound.ShowMessage += ShowMessage;
            this.Load += Form1_Load;
            this.FormClosing += MainForm_FormClosing;
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
                UpdateUI(() => excelPathTextBox.Text = openFileDialog.FileName);
            }
        }

        private async void StartButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(usernameTextBox.Text) || string.IsNullOrWhiteSpace(passwordTextBox.Text))
            {
                MessageBox.Show("Please enter both username and password before proceeding.", "Missing Credentials", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // For Neuware Buchen
            if (mode == "Neuware")
            {
                if (string.IsNullOrWhiteSpace(locationTextBox.Text) ||
                    string.IsNullOrWhiteSpace(jobOrPoTextBox.Text) ||
                    string.IsNullOrWhiteSpace(excelPathTextBox.Text))
                {
                    MessageBox.Show("Please fill out all fields for Neuware Buchen mode before proceeding.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            // For B2B Buchen
            else if (mode == "B2B")
            {
                if (string.IsNullOrWhiteSpace(locationTextBox.Text) ||
                    string.IsNullOrWhiteSpace(jobOrPoTextBox.Text) ||
                    string.IsNullOrWhiteSpace(excelPathTextBox.Text))
                {
                    MessageBox.Show("Please fill out all fields for B2B Buchen mode before proceeding.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            // For Manual Outbound
            else if (mode == "Manual_outbound")
            {
                if (string.IsNullOrWhiteSpace(CreditcomboBox.Text) ||
                    string.IsNullOrWhiteSpace(excelPathTextBox.Text))
                {
                    MessageBox.Show("Please fill out all fields for Manual Outbound mode before proceeding.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Please select a mode before proceeding.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Update the status label to show "Logging in"
            UpdateUI(() => statusLabel.Text = "Logging in");

            bool loginSuccess = await Task.Run(() =>
            {
                if (cancellationTokenSource.Token.IsCancellationRequested)
                {
                    return false; // Or an appropriate value to indicate that the task was cancelled
                }
                return global_functions.LoginToSite(excelPathTextBox.Text, usernameTextBox.Text, passwordTextBox.Text);
            }, cancellationTokenSource.Token);
            try
            {
                if (loginSuccess)
                {
                    //ShowMessage("Login successful");
                    UpdateUI(() => statusLabel.Text = "Logged in");
                    try
                    {
                        switch (mode)
                        {
                            case "Neuware":
                                await Task.Run(() =>
                                {
                                    if (cancellationTokenSource.Token.IsCancellationRequested)
                                    {
                                        return; // Exit if cancellation was requested
                                    }
                                    Neuware.start_neuware(locationTextBox.Text, jobOrPoTextBox.Text, cancellationTokenSource.Token);
                                });
                                break;
                            case "B2B":
                                await Task.Run(() =>
                                {
                                    if (cancellationTokenSource.Token.IsCancellationRequested)
                                    {
                                        return; // Exit if cancellation was requested
                                    }
                                    B2B.start_b2b(locationTextBox.Text, jobOrPoTextBox.Text, creditType, cancellationTokenSource.Token);
                                });
                                break;
                            case "Manual_outbound":
                                await Task.Run(() =>
                                {
                                    if (cancellationTokenSource.Token.IsCancellationRequested)
                                    {
                                        return; // Exit if cancellation was requested
                                    }
                                    manual_outbound.start_manual_outbound(cancellationTokenSource.Token);
                                });
                                break;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // This exception will be thrown if the task is cancelled.
                        UpdateUI(() => statusLabel.Text = "Operation was cancelled");
                        return;
                    }
                    // Process the Excel file after successful login
                    UpdateUI(() => statusLabel.Text = "Devices have been added");
                    ShowMessage("The devices have been booked, please check the excel file to confirm");
                }
                else
                {
                    ShowMessage("Login failed. Check username and password");
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
                excelPathTextBox.Text = "";
                locationTextBox.Text = "";
                jobOrPoTextBox.Text = "";
                if (comboBoxMode != null)
                    comboBoxMode.SelectedIndex = -1; // Assuming comboBox1 is the name of your ComboBox; adjust accordingly

                if (CreditcomboBox != null)
                    CreditcomboBox.SelectedIndex = -1;

                // Close the global_functions.driver and package
                if (global_functions.driver != null)
                {
                    global_functions.driver.Quit();
                    global_functions.driver = null;
                }

                if (global_functions.package != null)
                {
                    global_functions.package.Save();
                    global_functions.package.Dispose();
                    global_functions.package = null;
                }
            }
        }

        private void stopButton_Click(object sender, EventArgs e)
        {
            // 3. Signal tasks/threads to stop
            cancellationTokenSource.Cancel();

            // 2. Reset all the TextBox values
            locationTextBox.Text = "";
            jobOrPoTextBox.Text = "";
            usernameTextBox.Text = "";
            passwordTextBox.Text = "";
            excelPathTextBox.Text = "";
            // ... Add other TextBox controls as needed

            // Close the WebDriver if it exists
            if (global_functions.driver != null)
            {
                global_functions.driver.Quit();
                global_functions.driver = null;
            }

            // Close the Excel package if it exists
            if (global_functions.package != null)
            {
                global_functions.package.Save();
                global_functions.package.Dispose();
                global_functions.package = null;
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

        private void UpdateUI_statusLabel(string statusMessage)
        {
            UpdateUI(() => statusLabel.Text = statusMessage);
        }

        private void UpdateUI_progressBar(int value, string percentageText = "")
        {
            UpdateUI(() =>
            {
                progressBar.Value = value;
                percentLabel.Text = percentageText;
            });
        }

        private void UpdateUI_SetMaxProgressBar(int maxValue)
        {
            UpdateUI(() => progressBar.Maximum = maxValue);
        }

        public static void ShowMessage(string message)
        {
            MessageBox.Show(message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            // Bring the form to the front when the application starts
            this.BringToFront();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // Close the WebDriver if it exists
            if (global_functions.driver != null)
            {
                global_functions.driver.Quit();
                global_functions.driver = null;
            }

            // Close the Excel package if it exists
            if (global_functions.package != null)
            {
                global_functions.package.Dispose();
                global_functions.package = null;
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (About aboutBox = new About())
            {
                aboutBox.ShowDialog(this);
            }
        }

        private void LocationTextBox_Enter(object sender, EventArgs e)
        {
            toolTipLocation.Show("Please fill a sorting location", locationTextBox);
        }
        private void PoNoTextBox_Enter(object sender, EventArgs e)
        {
            toolTipJobOrPoNo.Show("Please fill a valid Pono or Job ID number", jobOrPoTextBox);
        }
        private void excelFilePathTextBox_Enter(object sender, EventArgs e)
        {
            toolTipExcel.Show("Columns should be:\nA:Serial number\nB:part number", excelPathTextBox);
        }

        private void createExcelTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                var ws = excel.Workbook.Worksheets.Add("Neuware");

                // Headers
                if (mode == "Manual_outbound")
                {
                    ws.Cells["A1"].Value = "Old ScanID";
                    ws.Cells["B1"].Value = "New Serial";
                }
                else
                {
                    ws.Cells["A1"].Value = "Serial number";
                    ws.Cells["B1"].Value = "Part number";
                }

                // Define the range for the table.
                var tableRange = ws.Cells["A1:B2"];

                // Create a table based on this range.
                var table = ws.Tables.Add(tableRange, "Table");

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
                saveFileDialog.FileName = "template_" + mode;
                saveFileDialog.DefaultExt = ".xlsx";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo fi = new FileInfo(saveFileDialog.FileName);
                    excel.SaveAs(fi);
                }
            }
        }

        private void comboBoxMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Reset visibility for all controls first
            locationLabel.Visible = false;
            locationTextBox.Visible = false;
            jobOrPoLabel.Visible = false;
            jobOrPoTextBox.Visible = false;
            browseButton.Visible = false;
            excelPathTextBox.Visible = false;
            creditLabel.Visible = false;
            CreditcomboBox.Visible = false;

            // Check the selected item in the combobox
            string selectedMode = comboBoxMode.SelectedItem?.ToString() ?? string.Empty;
            switch (selectedMode)
            {
                case "Neuware Buchen":
                    mode = "Neuware";
                    locationLabel.Visible = true;
                    locationTextBox.Visible = true;
                    jobOrPoLabel.Text = "PoNo";
                    jobOrPoLabel.Visible = true;
                    jobOrPoTextBox.Visible = true;
                    browseButton.Visible = true;
                    excelPathTextBox.Visible = true;
                    break;
                case "B2B Buchen":
                    mode = "B2B";
                    locationLabel.Visible = true;
                    locationTextBox.Visible = true;
                    jobOrPoLabel.Text = "Job ID";
                    jobOrPoLabel.Visible = true;
                    jobOrPoTextBox.Visible = true;
                    browseButton.Visible = true;
                    excelPathTextBox.Visible = true;
                    break;
                case "Manual Outbound":
                    mode = "Manual_outbound";
                    creditLabel.Visible = true;
                    CreditcomboBox.Visible = true;
                    browseButton.Visible = true;
                    excelPathTextBox.Visible = true;
                    break;
            }
        }

        private void CreditcomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedMode = CreditcomboBox.SelectedItem?.ToString() ?? string.Empty;
            switch (selectedMode)
            {
                case "Full Credit":
                    creditType = "Full Credit";
                    break;
                case "Swap":
                    creditType = "Swap";
                    break;
            }
        }

    }
}