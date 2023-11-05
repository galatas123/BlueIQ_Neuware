using OfficeOpenXml;
using System.Globalization;
using System.ComponentModel.DataAnnotations;
using System.Threading;

namespace BlueIQ_Neuware
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource cancellationTokenSource = new();

        public Form1()
        {
            InitializeComponent();
            Global_functions.StatusUpdated += UpdateUI_statusLabel;
            Neuware.ProgressUpdated += UpdateUI_progressBar;
            Neuware.StatusUpdated += UpdateUI_statusLabel;
            Neuware.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            Neuware.ShowMessage += ShowMessage;
            B2B.ProgressUpdated += UpdateUI_progressBar;
            B2B.StatusUpdated += UpdateUI_statusLabel;
            B2B.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            B2B.ShowMessage += ShowMessage;
            B2B.ShowMessageYesNo += ShowYesNoMessage;
            Manual_outbound.ProgressUpdated += UpdateUI_progressBar;
            Manual_outbound.StatusUpdated += UpdateUI_statusLabel;
            Manual_outbound.SetMaxProgress += UpdateUI_SetMaxProgressBar;
            Manual_outbound.ShowMessage += ShowMessage;
            Manual_outbound.RequestSaveFile += ShowSaveFileDialog;
            this.Load += Form1_Load;
            this.FormClosing += MainForm_FormClosing;
            SetLanguage(Global_functions.GetSettings());
        }

        private async void StartButton_Click(object sender, EventArgs e)
        {
            try
            {
                startButton.Enabled = false;
                if (string.IsNullOrWhiteSpace(usernameTextBox.Text) || string.IsNullOrWhiteSpace(passwordTextBox.Text))
                {
                    ShowMessage(Languages.Resources.FILL_FIELDS_MESS, MessageBoxIcon.Warning);
                    return;
                }
                // Use a switch statement to handle the different modes
                switch (JobInfo.Current.Mode)
                {
                    case "Neuware":
                    case "B2B": // Combine "Neuware" and "B2B" since they have similar validation requirements
                        if (string.IsNullOrWhiteSpace(locationTextBox.Text) ||
                            string.IsNullOrWhiteSpace(jobOrPoTextBox.Text) ||
                            string.IsNullOrWhiteSpace(excelPathTextBox.Text) ||
                            (JobInfo.Current.Mode == "B2B" && string.IsNullOrWhiteSpace(CreditcomboBox.Text))) // Additional check for "B2B"
                        {
                            ShowMessage(Languages.Resources.FILL_FIELDS_MESS, MessageBoxIcon.Warning);
                            return;
                        }
                        break;

                    case "Manual_outbound":
                        if (string.IsNullOrWhiteSpace(excelPathTextBox.Text))
                        {
                            ShowMessage(Languages.Resources.FILL_FIELDS_MESS, MessageBoxIcon.Warning);
                            return;
                        }
                        break;

                    default: // No valid mode selected
                        ShowMessage(Languages.Resources.FILL_FIELDS_MESS, MessageBoxIcon.Warning);
                        return;
                }

                //update variables
                JobInfo.Current.Location = locationTextBox.Text;
                JobInfo.Current.JobOrPoNO = jobOrPoTextBox.Text;

                // Update the status label to show "Logging in"
                UpdateUI_statusLabel(Languages.Resources.LOGGING);
                bool loginSuccess = await Task.Run(() =>
                {
                    if (cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        return false; // Or an appropriate value to indicate that the task was cancelled
                    }
                    return Global_functions.LoginToSite(excelPathTextBox.Text, usernameTextBox.Text, passwordTextBox.Text);
                }, cancellationTokenSource.Token);

                if (loginSuccess)
                {
                    //ShowMessage("Login successful");
                    UpdateUI_statusLabel(Languages.Resources.LOGGED);
                    try
                    {
                        if (!cancellationTokenSource.Token.IsCancellationRequested)
                        {
                            await Task.Run(() =>
                            {
                                switch (JobInfo.Current.Mode)
                                {
                                    case "Neuware":
                                        Neuware.Start_neuware(cancellationTokenSource.Token);
                                        break;
                                    case "B2B":
                                        B2B.Start_b2b(cancellationTokenSource.Token);
                                        break;
                                    case "Manual_outbound":
                                        Manual_outbound.Start_manual_outbound(cancellationTokenSource.Token);
                                        break;
                                    default:
                                        throw new InvalidOperationException("Invalid mode");
                                }
                            }, cancellationTokenSource.Token);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        UpdateUI_statusLabel("Operation was cancelled");
                    }
                    catch (InvalidOperationException ex)
                    {
                        ShowMessage("Invalid mode: " + ex.Message, MessageBoxIcon.Warning);
                    }
                    // Process the Excel file after successful login
                    UpdateUI_statusLabel(Languages.Resources.BOOKED_ALL_STATUS);
                    ShowMessage(Languages.Resources.BOOKED_ALL_MESS);
                }
                else
                {
                    if (!cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        ShowMessage(Languages.Resources.LOG_FAIL_MESS, MessageBoxIcon.Error);
                        UpdateUI_statusLabel(Languages.Resources.LOG_FAIL_STATUS);
                    }
                    else
                    {
                        UpdateUI_statusLabel(Languages.Resources.PROGRAM_STOPPED);
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle the exception
                ShowMessage(Languages.Resources.ERROR_GEN_MESS + ex.Message, MessageBoxIcon.Error);
                UpdateUI_statusLabel(Languages.Resources.ERROR_GEN_STATUS);
            }
            finally
            {
                CleanUp();
                UpdateUI_statusLabel(Languages.Resources.PRO_FINISHED);
                startButton.Enabled = true;
            }
        }


        // User Interface Update Methods
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

        public static void ShowMessage(string message, MessageBoxIcon icon = MessageBoxIcon.Information)
        {
            MessageBox.Show(message, "Message", MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
        }

        public static bool ShowYesNoMessage(string message)
        {
            DialogResult result = MessageBox.Show(message, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            return result == DialogResult.Yes;
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


        // file path and excel Methods
        private void ShowSaveFileDialog(out string? savedFilePath)
        {
            string? resultPath = null;

            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    resultPath = ShowSaveFileDialogActual();
                });
            }
            else
            {
                resultPath = ShowSaveFileDialogActual();
            }

            savedFilePath = resultPath;
        }

        private static string? ShowSaveFileDialogActual()
        {
            SaveFileDialog saveFileDialog = new()
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                return saveFileDialog.FileName;
            }
            return null;
        }

        private void CreateExcelTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // for free version
            using ExcelPackage excel = new();
            var ws = excel.Workbook.Worksheets.Add("Neuware");

            // Headers
            if (JobInfo.Current.Mode == "Manual_outbound")
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
            table.TableStyle = OfficeOpenXml.Table.TableStyles.Light9;// Start with no style
            table.ShowHeader = true;
            table.ShowFilter = true;

            // Auto fit columns
            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            SaveFileDialog saveFileDialog = new()
            {
                FileName = "template_" + JobInfo.Current.Mode,
                DefaultExt = ".xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo fi = new(saveFileDialog.FileName);
                excel.SaveAs(fi);
            }
        }


        //Element interface methods
        private void StopButton_Click(object sender, EventArgs e)
        {
            CleanUp();
            ShowMessage(Languages.Resources.APP_STOPPED);
            UpdateUI(() => statusLabel.Text = Languages.Resources.PROGRAM_STOPPED);
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

        private void EnglishToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("en");
        }

        private void GermanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetLanguage("de");
        }

        private void SetLanguage(string cultureName)
        {
            Global_functions.UpdateLanguageInSettingsFile(cultureName);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(cultureName);
            // Now refresh the current form or reload it to see the changes
            this.Controls.Clear();
            this.InitializeComponent();
            // Any other initialization code, like setting up menu items, etc.
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using About aboutBox = new();
            aboutBox.ShowDialog(this);
        }

        private void LocationTextBox_Enter(object sender, EventArgs e)
        {
            toolTipLocation.Show(Languages.Resources.LOCATION_TIP, locationTextBox);
        }

        private void PoNoTextBox_Enter(object sender, EventArgs e)
        {
            toolTipJobOrPoNo.Show(Languages.Resources.PONO_TIP, jobOrPoTextBox);
        }

        private void ExcelFilePathTextBox_Enter(object sender, EventArgs e)
        {
            if (JobInfo.Current.Mode == "Manual_outbound")
                toolTipExcel.Show(Languages.Resources.EXCEL_FILE_TIP2, excelPathTextBox);
            else
                toolTipExcel.Show(Languages.Resources.EXCEL_FILE_TIP1, excelPathTextBox);
        }

        private void ComboBoxMode_SelectedIndexChanged(object sender, EventArgs e)
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
                    JobInfo.Current.Mode = "Neuware";
                    locationLabel.Visible = true;
                    locationTextBox.Visible = true;
                    jobOrPoLabel.Text = "PoNo";
                    jobOrPoLabel.Visible = true;
                    jobOrPoTextBox.Visible = true;
                    browseButton.Visible = true;
                    excelPathTextBox.Visible = true;
                    break;

                case "B2B Buchen":
                    JobInfo.Current.Mode = "B2B";
                    locationLabel.Visible = true;
                    locationTextBox.Visible = true;
                    jobOrPoLabel.Text = "Job ID";
                    jobOrPoLabel.Visible = true;
                    jobOrPoTextBox.Visible = true;
                    browseButton.Visible = true;
                    excelPathTextBox.Visible = true;
                    creditLabel.Visible = true;
                    CreditcomboBox.Visible = true;
                    break;

                case "Manual Outbound":
                    JobInfo.Current.Mode = "Manual_outbound";
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
                    JobInfo.Current.CreditType = "Full Credit";
                    break;

                case "Swap":
                    JobInfo.Current.CreditType = "Swap";
                    break;
            }
        }

        //Form methods
        private void Form1_Load(object? sender, EventArgs e)
        {
            // Bring the form to the front when the application starts
            this.BringToFront();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            CleanUp();
        }



        private void CleanUp()
        {
            // 1. Signal tasks/threads to stop
            try
            {
                if (cancellationTokenSource != null)
                {
                    cancellationTokenSource.Cancel();
                    Thread.Sleep(10);
                    cancellationTokenSource.Dispose();
                    cancellationTokenSource = new CancellationTokenSource();
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
            }

            // 2. Close the WebDriver if it exists
            try
            {
                if (Global_functions.driver != null)
                {
                    Global_functions.driver.Quit();
                    Global_functions.driver = null;
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
            }

            // 3. Close the Excel package if it exists
            try
            {
                if (Global_functions.package != null)
                {
                    if (Global_functions.package.Workbook.Worksheets.Count > 0)
                    {
                        Global_functions.package.Save();
                    }
                    Global_functions.package.Dispose();
                    Global_functions.package = null;
                }
            }
            catch (Exception ex)
            {
                Global_functions.LogError(Global_functions.GetCallerFunctionName(), ex.ToString());
            }

            locationTextBox.Text = "";
            jobOrPoTextBox.Text = "";
            usernameTextBox.Text = "";
            passwordTextBox.Text = "";
            excelPathTextBox.Text = "";
            JobInfo.Current.Mode = "";
            JobInfo.Current.CreditType = "";
            JobInfo.Current.LoadId = "";
            JobInfo.Current.JobOrPoNO = "";
            JobInfo.Current.PalletId = "";
            JobInfo.Current.Location = "";

            if (comboBoxMode != null)
                comboBoxMode.SelectedIndex = -1;
            if (CreditcomboBox != null)
                CreditcomboBox.SelectedIndex = -1;

            locationLabel.Visible = false;
            locationTextBox.Visible = false;
            jobOrPoLabel.Visible = false;
            jobOrPoTextBox.Visible = false;
            browseButton.Visible = false;
            excelPathTextBox.Visible = false;
            creditLabel.Visible = false;
            CreditcomboBox.Visible = false;
        }
    }


    public static class JobInfo
    {
        public static JobDetails Current { get; set; } = new JobDetails();

        public class JobDetails
        {
            public string? PalletId { get; set; }
            public string? LoadId { get; set; }
            public string? JobOrPoNO { get; set; }
            public string? Location { get; set; }
            public string? CreditType { get; set; }
            public string? Mode { get; set; }
        }
    }
}