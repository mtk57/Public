using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleExcelDiff
{
    public partial class MainForm : Form
    {
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();

            settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleExcelDiff_Settings.json");

            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;

            lblStatus.Text = "準備完了";

            this.txtPathSrc.AllowDrop = true;
            this.txtPathSrc.DragEnter += new DragEventHandler(txtDirSrc_DragEnter);
            this.txtPathSrc.DragDrop += new DragEventHandler(txtDirSrc_DragDrop);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version}";

            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void LoadSettings()
        {
            // TBD
        }

        private void SaveSettings()
        {
            // TBD
        }

        private void txtDirSrc_DragEnter(object sender, DragEventArgs e)
        {
            // TBD
        }

        private void txtDirSrc_DragDrop(object sender, DragEventArgs e)
        {
            // TBD
        }


        private void btnBrowseSrc_Click(object sender, EventArgs e)
        {
            // TBD
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            // TBD
        }
    }


    [DataContract]
    internal class SheetMergeToolSettings
    {
            // TBD
    }
}