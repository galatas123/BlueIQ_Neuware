namespace BlueIQ_Neuware
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            usernameTextBox = new TextBox();
            passwordTextBox = new TextBox();
            username_label = new Label();
            password_label = new Label();
            excelPathTextBox = new TextBox();
            progressBar = new ProgressBar();
            statusStrip1 = new StatusStrip();
            statusLabel = new ToolStripStatusLabel();
            toolStripStatusLabel1 = new ToolStripStatusLabel();
            startButton = new Button();
            browseButton = new Button();
            locationLabel = new Label();
            locationTextBox = new TextBox();
            toolTipExcel = new ToolTip(components);
            toolStrip1 = new ToolStrip();
            toolStripDropDownButton1 = new ToolStripDropDownButton();
            aboutToolStripMenuItem = new ToolStripMenuItem();
            createExcelTemplateToolStripMenuItem = new ToolStripMenuItem();
            percentLabel = new Label();
            jobOrPoLabel = new Label();
            jobOrPoTextBox = new TextBox();
            toolTipLocation = new ToolTip(components);
            toolTipJobOrPoNo = new ToolTip(components);
            comboBoxMode = new ComboBox();
            taskLabel = new Label();
            CreditcomboBox = new ComboBox();
            creditLabel = new Label();
            stopButton = new Button();
            statusStrip1.SuspendLayout();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // usernameTextBox
            // 
            usernameTextBox.BackColor = Color.White;
            usernameTextBox.BorderStyle = BorderStyle.None;
            usernameTextBox.Location = new Point(387, 168);
            usernameTextBox.Name = "usernameTextBox";
            usernameTextBox.Size = new Size(100, 16);
            usernameTextBox.TabIndex = 0;
            // 
            // passwordTextBox
            // 
            passwordTextBox.BackColor = Color.White;
            passwordTextBox.BorderStyle = BorderStyle.None;
            passwordTextBox.Location = new Point(387, 235);
            passwordTextBox.Name = "passwordTextBox";
            passwordTextBox.PasswordChar = '*';
            passwordTextBox.Size = new Size(100, 16);
            passwordTextBox.TabIndex = 1;
            // 
            // username_label
            // 
            username_label.AutoSize = true;
            username_label.ForeColor = SystemColors.ActiveCaptionText;
            username_label.Location = new Point(307, 169);
            username_label.Name = "username_label";
            username_label.Size = new Size(60, 15);
            username_label.TabIndex = 2;
            username_label.Text = "Username";
            // 
            // password_label
            // 
            password_label.AutoSize = true;
            password_label.Location = new Point(307, 236);
            password_label.Name = "password_label";
            password_label.Size = new Size(57, 15);
            password_label.TabIndex = 3;
            password_label.Text = "Password";
            password_label.TextAlign = ContentAlignment.TopRight;
            // 
            // excelPathTextBox
            // 
            excelPathTextBox.BackColor = Color.White;
            excelPathTextBox.Location = new Point(83, 102);
            excelPathTextBox.Name = "excelPathTextBox";
            excelPathTextBox.Size = new Size(100, 23);
            excelPathTextBox.TabIndex = 4;
            excelPathTextBox.Visible = false;
            excelPathTextBox.Enter += excelFilePathTextBox_Enter;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(2, 393);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(763, 23);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.TabIndex = 5;
            // 
            // statusStrip1
            // 
            statusStrip1.BackColor = SystemColors.Highlight;
            statusStrip1.Items.AddRange(new ToolStripItem[] { statusLabel, toolStripStatusLabel1 });
            statusStrip1.Location = new Point(0, 428);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(800, 22);
            statusStrip1.TabIndex = 6;
            statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = false;
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(150, 17);
            statusLabel.Text = "status";
            // 
            // toolStripStatusLabel1
            // 
            toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            toolStripStatusLabel1.Size = new Size(635, 17);
            toolStripStatusLabel1.Spring = true;
            toolStripStatusLabel1.Text = "© 2023 Ingram Micro Services";
            // 
            // startButton
            // 
            startButton.BackColor = SystemColors.Highlight;
            startButton.FlatStyle = FlatStyle.Flat;
            startButton.ForeColor = Color.Black;
            startButton.Location = new Point(283, 314);
            startButton.Name = "startButton";
            startButton.Size = new Size(126, 23);
            startButton.TabIndex = 7;
            startButton.Text = "start";
            startButton.UseVisualStyleBackColor = false;
            startButton.Click += StartButton_Click;
            // 
            // browseButton
            // 
            browseButton.Location = new Point(2, 102);
            browseButton.Name = "browseButton";
            browseButton.Size = new Size(75, 23);
            browseButton.TabIndex = 8;
            browseButton.Text = "browse...";
            browseButton.UseVisualStyleBackColor = true;
            browseButton.Visible = false;
            browseButton.Click += BrowseButton_Click;
            // 
            // locationLabel
            // 
            locationLabel.AutoSize = true;
            locationLabel.Location = new Point(12, 29);
            locationLabel.Name = "locationLabel";
            locationLabel.Size = new Size(53, 15);
            locationLabel.TabIndex = 9;
            locationLabel.Text = "Location";
            locationLabel.Visible = false;
            // 
            // locationTextBox
            // 
            locationTextBox.BackColor = Color.White;
            locationTextBox.BorderStyle = BorderStyle.None;
            locationTextBox.Location = new Point(83, 28);
            locationTextBox.Name = "locationTextBox";
            locationTextBox.Size = new Size(100, 16);
            locationTextBox.TabIndex = 10;
            locationTextBox.Visible = false;
            locationTextBox.Enter += LocationTextBox_Enter;
            // 
            // toolStrip1
            // 
            toolStrip1.BackColor = SystemColors.Highlight;
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripDropDownButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(800, 25);
            toolStrip1.TabIndex = 12;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripDropDownButton1
            // 
            toolStripDropDownButton1.DisplayStyle = ToolStripItemDisplayStyle.Text;
            toolStripDropDownButton1.DropDownItems.AddRange(new ToolStripItem[] { aboutToolStripMenuItem, createExcelTemplateToolStripMenuItem });
            toolStripDropDownButton1.Image = (Image)resources.GetObject("toolStripDropDownButton1.Image");
            toolStripDropDownButton1.ImageTransparentColor = Color.Magenta;
            toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            toolStripDropDownButton1.Size = new Size(51, 22);
            toolStripDropDownButton1.Text = "Menu";
            // 
            // aboutToolStripMenuItem
            // 
            aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            aboutToolStripMenuItem.Size = new Size(187, 22);
            aboutToolStripMenuItem.Text = "About";
            aboutToolStripMenuItem.Click += aboutToolStripMenuItem_Click;
            // 
            // createExcelTemplateToolStripMenuItem
            // 
            createExcelTemplateToolStripMenuItem.Name = "createExcelTemplateToolStripMenuItem";
            createExcelTemplateToolStripMenuItem.Size = new Size(187, 22);
            createExcelTemplateToolStripMenuItem.Text = "create Excel Template";
            createExcelTemplateToolStripMenuItem.Click += createExcelTemplateToolStripMenuItem_Click;
            // 
            // percentLabel
            // 
            percentLabel.AutoSize = true;
            percentLabel.Location = new Point(771, 401);
            percentLabel.Name = "percentLabel";
            percentLabel.Size = new Size(23, 15);
            percentLabel.TabIndex = 13;
            percentLabel.Text = "0%";
            // 
            // jobOrPoLabel
            // 
            jobOrPoLabel.AutoSize = true;
            jobOrPoLabel.Location = new Point(12, 68);
            jobOrPoLabel.Name = "jobOrPoLabel";
            jobOrPoLabel.Size = new Size(37, 15);
            jobOrPoLabel.TabIndex = 14;
            jobOrPoLabel.Text = "PoNo";
            jobOrPoLabel.Visible = false;
            // 
            // jobOrPoTextBox
            // 
            jobOrPoTextBox.Location = new Point(83, 60);
            jobOrPoTextBox.Name = "jobOrPoTextBox";
            jobOrPoTextBox.Size = new Size(100, 23);
            jobOrPoTextBox.TabIndex = 15;
            jobOrPoTextBox.Visible = false;
            jobOrPoTextBox.Enter += PoNoTextBox_Enter;
            // 
            // comboBoxMode
            // 
            comboBoxMode.FormattingEnabled = true;
            comboBoxMode.Items.AddRange(new object[] { "Neuware Buchen", "B2B Buchen", "Manual Outbound" });
            comboBoxMode.Location = new Point(387, 102);
            comboBoxMode.Name = "comboBoxMode";
            comboBoxMode.Size = new Size(121, 23);
            comboBoxMode.TabIndex = 16;
            comboBoxMode.SelectedIndexChanged += comboBoxMode_SelectedIndexChanged;
            // 
            // taskLabel
            // 
            taskLabel.AutoSize = true;
            taskLabel.Location = new Point(296, 106);
            taskLabel.Name = "taskLabel";
            taskLabel.Size = new Size(71, 15);
            taskLabel.TabIndex = 17;
            taskLabel.Text = "Choose task";
            // 
            // CreditcomboBox
            // 
            CreditcomboBox.FormattingEnabled = true;
            CreditcomboBox.Items.AddRange(new object[] { "Full Credit", "Swap" });
            CreditcomboBox.Location = new Point(74, 146);
            CreditcomboBox.Name = "CreditcomboBox";
            CreditcomboBox.Size = new Size(121, 23);
            CreditcomboBox.TabIndex = 18;
            CreditcomboBox.Visible = false;
            CreditcomboBox.SelectedIndexChanged += CreditcomboBox_SelectedIndexChanged;
            // 
            // creditLabel
            // 
            creditLabel.AutoSize = true;
            creditLabel.Location = new Point(2, 149);
            creditLabel.Name = "creditLabel";
            creditLabel.Size = new Size(66, 15);
            creditLabel.TabIndex = 19;
            creditLabel.Text = "Credit Type";
            creditLabel.Visible = false;
            // 
            // stopButton
            // 
            stopButton.BackColor = SystemColors.Highlight;
            stopButton.FlatStyle = FlatStyle.Flat;
            stopButton.ForeColor = Color.Black;
            stopButton.Location = new Point(474, 314);
            stopButton.Name = "stopButton";
            stopButton.Size = new Size(126, 23);
            stopButton.TabIndex = 20;
            stopButton.Text = "stop";
            stopButton.UseVisualStyleBackColor = false;
            stopButton.Click += stopButton_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.LightGray;
            ClientSize = new Size(800, 450);
            Controls.Add(stopButton);
            Controls.Add(creditLabel);
            Controls.Add(CreditcomboBox);
            Controls.Add(taskLabel);
            Controls.Add(comboBoxMode);
            Controls.Add(jobOrPoTextBox);
            Controls.Add(jobOrPoLabel);
            Controls.Add(percentLabel);
            Controls.Add(toolStrip1);
            Controls.Add(locationTextBox);
            Controls.Add(locationLabel);
            Controls.Add(browseButton);
            Controls.Add(startButton);
            Controls.Add(statusStrip1);
            Controls.Add(progressBar);
            Controls.Add(excelPathTextBox);
            Controls.Add(password_label);
            Controls.Add(username_label);
            Controls.Add(passwordTextBox);
            Controls.Add(usernameTextBox);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Blue IQ Neuware";
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox usernameTextBox;
        private TextBox passwordTextBox;
        private Label username_label;
        private Label password_label;
        private TextBox excelPathTextBox;
        private ProgressBar progressBar;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel statusLabel;
        private Button startButton;
        private Button browseButton;
        private Label locationLabel;
        private TextBox locationTextBox;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolTip toolTipExcel;
        private ToolStrip toolStrip1;
        private ToolStripDropDownButton toolStripDropDownButton1;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private ToolStripMenuItem createExcelTemplateToolStripMenuItem;
        private Label percentLabel;
        private Label jobOrPoLabel;
        private TextBox jobOrPoTextBox;
        private ToolTip toolTipLocation;
        private ToolTip toolTipJobOrPoNo;
        private ComboBox comboBoxMode;
        private Label taskLabel;
        private ComboBox CreditcomboBox;
        private Label creditLabel;
        private Button stopButton;
    }
}