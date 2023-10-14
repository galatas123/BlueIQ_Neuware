namespace BlueIQ_Neuware
{
    partial class About
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(About));
            tableLayoutPanel = new TableLayoutPanel();
            labelProductName = new Label();
            labelVersion = new Label();
            labelCopyright = new Label();
            labelCompanyName = new Label();
            logoPictureBox = new PictureBox();
            okButton = new Button();
            textBoxDescription = new TextBox();
            tableLayoutPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)logoPictureBox).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel
            // 
            tableLayoutPanel.ColumnCount = 2;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48.1144333F));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 51.8855667F));
            tableLayoutPanel.Controls.Add(labelProductName, 1, 0);
            tableLayoutPanel.Controls.Add(labelVersion, 1, 1);
            tableLayoutPanel.Controls.Add(labelCopyright, 1, 2);
            tableLayoutPanel.Controls.Add(labelCompanyName, 1, 3);
            tableLayoutPanel.Controls.Add(logoPictureBox, 0, 0);
            tableLayoutPanel.Controls.Add(okButton, 1, 6);
            tableLayoutPanel.Controls.Add(textBoxDescription, 1, 4);
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.Location = new Point(10, 10);
            tableLayoutPanel.Margin = new Padding(4, 3, 4, 3);
            tableLayoutPanel.Name = "tableLayoutPanel";
            tableLayoutPanel.RowCount = 7;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10.03924F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10.0392418F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10.0392418F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 32.2981377F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 27.95031F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10.03924F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36F));
            tableLayoutPanel.Size = new Size(523, 314);
            tableLayoutPanel.TabIndex = 0;
            // 
            // labelProductName
            // 
            labelProductName.Dock = DockStyle.Fill;
            labelProductName.Location = new Point(258, 0);
            labelProductName.Margin = new Padding(7, 0, 4, 0);
            labelProductName.MaximumSize = new Size(0, 20);
            labelProductName.Name = "labelProductName";
            labelProductName.Size = new Size(261, 20);
            labelProductName.TabIndex = 19;
            labelProductName.Text = "BlueIQ Neuware Buchung";
            labelProductName.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // labelVersion
            // 
            labelVersion.Dock = DockStyle.Fill;
            labelVersion.Location = new Point(258, 27);
            labelVersion.Margin = new Padding(7, 0, 4, 0);
            labelVersion.MaximumSize = new Size(0, 20);
            labelVersion.Name = "labelVersion";
            labelVersion.Size = new Size(261, 20);
            labelVersion.TabIndex = 0;
            labelVersion.Text = "Version: 1.0";
            labelVersion.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // labelCopyright
            // 
            labelCopyright.Dock = DockStyle.Fill;
            labelCopyright.Location = new Point(258, 54);
            labelCopyright.Margin = new Padding(7, 0, 4, 0);
            labelCopyright.MaximumSize = new Size(0, 20);
            labelCopyright.Name = "labelCopyright";
            labelCopyright.Size = new Size(261, 20);
            labelCopyright.TabIndex = 21;
            labelCopyright.Text = "© 2023 Ingram Micro Services";
            labelCopyright.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // labelCompanyName
            // 
            labelCompanyName.Dock = DockStyle.Fill;
            labelCompanyName.Location = new Point(258, 81);
            labelCompanyName.Margin = new Padding(7, 0, 4, 0);
            labelCompanyName.MaximumSize = new Size(0, 20);
            labelCompanyName.Name = "labelCompanyName";
            labelCompanyName.Size = new Size(261, 20);
            labelCompanyName.TabIndex = 22;
            labelCompanyName.Text = "Programmer: Anil Chikmet Oglou";
            labelCompanyName.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // logoPictureBox
            // 
            logoPictureBox.Dock = DockStyle.Fill;
            logoPictureBox.Image = (Image)resources.GetObject("logoPictureBox.Image");
            logoPictureBox.Location = new Point(4, 3);
            logoPictureBox.Margin = new Padding(4, 3, 4, 3);
            logoPictureBox.Name = "logoPictureBox";
            tableLayoutPanel.SetRowSpan(logoPictureBox, 6);
            logoPictureBox.Size = new Size(243, 268);
            logoPictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            logoPictureBox.TabIndex = 12;
            logoPictureBox.TabStop = false;
            // 
            // okButton
            // 
            okButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            okButton.DialogResult = DialogResult.Cancel;
            okButton.Location = new Point(431, 282);
            okButton.Margin = new Padding(4, 3, 4, 3);
            okButton.Name = "okButton";
            okButton.Size = new Size(88, 29);
            okButton.TabIndex = 24;
            okButton.Text = "&OK";
            // 
            // textBoxDescription
            // 
            textBoxDescription.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            textBoxDescription.Location = new Point(258, 173);
            textBoxDescription.Margin = new Padding(7, 3, 4, 3);
            textBoxDescription.Multiline = true;
            textBoxDescription.Name = "textBoxDescription";
            textBoxDescription.ReadOnly = true;
            textBoxDescription.ScrollBars = ScrollBars.Both;
            textBoxDescription.Size = new Size(261, 71);
            textBoxDescription.TabIndex = 23;
            textBoxDescription.TabStop = false;
            textBoxDescription.Text = "A programm for booking new devices to blue IQ.\r\nFor questions or requests please contact\r\nAnil.ChikmetOglou@ingrammicro.com";
            // 
            // About
            // 
            AcceptButton = okButton;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(543, 334);
            Controls.Add(tableLayoutPanel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Margin = new Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "About";
            Padding = new Padding(10);
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "About";
            tableLayoutPanel.ResumeLayout(false);
            tableLayoutPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)logoPictureBox).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel;
        private PictureBox logoPictureBox;
        private Label labelProductName;
        private Label labelVersion;
        private Label labelCopyright;
        private Label labelCompanyName;
        private TextBox textBoxDescription;
        private Button okButton;
    }
}
