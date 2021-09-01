using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoNotifier.BusinessLogic;

namespace AutoNotifier
{
    public partial class Form1 : Form
    {
        private readonly OutstandingDataManager _outstandingDataManager;

        public Form1(OutstandingDataManager outstandingDataManager)
        {
            _outstandingDataManager = outstandingDataManager;
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"Excel Files|*.xls;*.xlsx;*.xlsb";
            openFileDialog.FileOk += FileDialog_FileOk;
            openFileDialog.ShowDialog();
        }

        private void FileDialog_FileOk(object sender, CancelEventArgs e)
        {
            var fileDialog = (OpenFileDialog)sender;
            fileNameTextBox.Text = fileDialog.FileName;
            fileDialog.Dispose();
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            string fileName = fileNameTextBox.Text;
            string sheetName = sheetNameTextBox.Text;
            if (!string.IsNullOrEmpty(fileName) && !string.IsNullOrEmpty(sheetName))
            {
                BeginInvoke(new Action(() =>
                    messageLabel.Text = $"Reading outstanding file.\nPlease wait or you can minimize the window."));

                var outstandings = await Task.Run(() =>
                    _outstandingDataManager.GetOutstandingEnumerable(fileName, sheetName, null));

                var customerWiseOutstandings = await Task.Run(() =>
                    _outstandingDataManager.GetCustomerWiseOutstanding(outstandings));
            }
        }
    }
}
