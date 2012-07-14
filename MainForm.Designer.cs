namespace EskomParser
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.pdfFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.xlsFile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.progressOutput = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.quitButton = new System.Windows.Forms.Button();
            this.createXLSButton = new System.Windows.Forms.Button();
            this.loadPDFButton = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(2, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Billing PDF file";
            // 
            // pdfFile
            // 
            this.pdfFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.pdfFile.Location = new System.Drawing.Point(5, 101);
            this.pdfFile.Name = "pdfFile";
            this.pdfFile.ReadOnly = true;
            this.pdfFile.Size = new System.Drawing.Size(572, 20);
            this.pdfFile.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(2, 124);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "Billing Template";
            // 
            // xlsFile
            // 
            this.xlsFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.xlsFile.Location = new System.Drawing.Point(5, 143);
            this.xlsFile.Name = "xlsFile";
            this.xlsFile.ReadOnly = true;
            this.xlsFile.Size = new System.Drawing.Size(572, 20);
            this.xlsFile.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(2, 166);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(150, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Parsing progress output";
            // 
            // progressOutput
            // 
            this.progressOutput.BackColor = System.Drawing.SystemColors.Window;
            this.progressOutput.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.progressOutput.Location = new System.Drawing.Point(5, 185);
            this.progressOutput.Multiline = true;
            this.progressOutput.Name = "progressOutput";
            this.progressOutput.ReadOnly = true;
            this.progressOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.progressOutput.Size = new System.Drawing.Size(572, 282);
            this.progressOutput.TabIndex = 7;
            this.progressOutput.WordWrap = false;
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "pdf";
            this.openFileDialog.Filter = "Billing | *.pdf|Excel|*.xlsx";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "xls";
            this.saveFileDialog.Filter = "Excel | *.xlsx";
            // 
            // quitButton
            // 
            this.quitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.quitButton.Image = global::EskomParser.Properties.Resources.poweroff_icon;
            this.quitButton.Location = new System.Drawing.Point(509, 5);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(68, 68);
            this.quitButton.TabIndex = 8;
            this.toolTip.SetToolTip(this.quitButton, "Quit Application");
            this.quitButton.UseVisualStyleBackColor = true;
            this.quitButton.Click += new System.EventHandler(this.QuitButtonClick);
            // 
            // createXLSButton
            // 
            this.createXLSButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.createXLSButton.Enabled = false;
            this.createXLSButton.Image = global::EskomParser.Properties.Resources.excel_icon;
            this.createXLSButton.Location = new System.Drawing.Point(76, 5);
            this.createXLSButton.Name = "createXLSButton";
            this.createXLSButton.Size = new System.Drawing.Size(68, 68);
            this.createXLSButton.TabIndex = 1;
            this.toolTip.SetToolTip(this.createXLSButton, "Load Excel Template and Parse");
            this.createXLSButton.UseVisualStyleBackColor = true;
            this.createXLSButton.Click += new System.EventHandler(this.CreateXlsButtonClick);
            // 
            // loadPDFButton
            // 
            this.loadPDFButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.loadPDFButton.Image = global::EskomParser.Properties.Resources.pdf_icon;
            this.loadPDFButton.Location = new System.Drawing.Point(5, 5);
            this.loadPDFButton.Name = "loadPDFButton";
            this.loadPDFButton.Size = new System.Drawing.Size(68, 68);
            this.loadPDFButton.TabIndex = 0;
            this.toolTip.SetToolTip(this.loadPDFButton, "Load PDF Document");
            this.loadPDFButton.UseVisualStyleBackColor = true;
            this.loadPDFButton.Click += new System.EventHandler(this.LoadPdfButtonClick);
            // 
            // toolTip
            // 
            this.toolTip.IsBalloon = true;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Window;
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.loadPDFButton);
            this.panel1.Controls.Add(this.quitButton);
            this.panel1.Controls.Add(this.createXLSButton);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(583, 79);
            this.panel1.TabIndex = 9;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 78);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(583, 1);
            this.panel2.TabIndex = 9;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(583, 474);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.progressOutput);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.xlsFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.pdfFile);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Eskom Parser";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button loadPDFButton;
        private System.Windows.Forms.Button createXLSButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox pdfFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox xlsFile;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox progressOutput;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
    }
}

