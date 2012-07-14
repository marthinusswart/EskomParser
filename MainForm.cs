using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;


namespace EskomParser
{
    public partial class MainForm : Form
    {
        #region Fields

        private Thread _thread;

        #endregion

        #region Constructors

        public MainForm()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        public string GetSaveFile()
        {
            var file = string.Empty;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                file = saveFileDialog.FileName;
            }
            return file;
        }

        private void LoadPDFBill()
        {
            openFileDialog.FilterIndex = 1;
            openFileDialog.FileName = string.Empty;
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                pdfFile.Text = openFileDialog.FileName;
                createXLSButton.Enabled = true;
            }
        }

        private void LoadTemplate()
        {
            openFileDialog.FilterIndex = 2;
            openFileDialog.FileName = string.Empty;
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                xlsFile.Text = openFileDialog.FileName;
                _thread = new Thread(new ThreadStart(Parse));
                _thread.Start();
            }
        }

        private void Parse()
        {
            var parser = new PDF2XLSParser();
            parser.ProgressOutput = progressOutput;
            parser.MainForm = this;
            parser.Parse(pdfFile.Text, xlsFile.Text);
        }

        private void CreateXlsButtonClick(object sender, EventArgs e)
        {
            LoadTemplate();
        }

        private void LoadPdfButtonClick(object sender, EventArgs e)
        {
            LoadPDFBill();
        }

        private void QuitButtonClick(object sender, EventArgs e)
        {
            if ((_thread == null) || (!_thread.IsAlive))
            {
                Application.Exit();
            }
        }

        #endregion
    }
}
