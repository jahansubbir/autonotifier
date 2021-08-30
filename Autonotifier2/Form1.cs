using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autonotifier2.BusinessLogic;
using Autonotifier2.Models;
using Autonotifier2.Utilities;

namespace Autonotifier2
{
    public partial class Form1 : Form
    {
        private readonly OutstandingDataManager _outstandingDataManager;
        /*private readonly IExcelGateway _excelGateway;
        private readonly JsonColumnHeaderReader _columnHeaderReader;*/

        public Form1(/*IExcelGateway excelGateway, JsonColumnHeaderReader columnHeaderReader*/ OutstandingDataManager outstandingDataManager)
        {
            _outstandingDataManager = outstandingDataManager;
            
            InitializeComponent();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"Excel Files|*.xls;*.xlsx;*.xlsb";
            openFileDialog.FileOk += OpenFileDialog_FileOk;
            openFileDialog.ShowDialog();


        }

        private void OpenFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            var fileDialog = (OpenFileDialog)sender;
            filePathTextBox.Text = fileDialog.FileName;
            fileDialog.Dispose();
        }

        private async void CreateButton_Click(object sender, EventArgs e)
        {
            string filePath = filePathTextBox.Text;
            string sheetName = sheetNameTextBox.Text;

            if (!string.IsNullOrEmpty(filePath) && !string.IsNullOrEmpty(sheetName))
            {
                BeginInvoke(new Action(() =>
                    messageLabel.Text = $"Reading outstanding file.\nPlease wait or you can minimize the window."));

                var outstandings = await Task.Run(() =>
                    _outstandingDataManager.GetCustomerWiseOutstanding(filePath, sheetName, null));
                


            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
