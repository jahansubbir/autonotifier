using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autonotifier2.BusinessLogic;
using Autonotifier2.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Autonotifier2.DataAccess
{
    class NpoiReport : NpoiExcel, IReporting
    {
        public void CreateReport(string fileName, IEnumerable<Outstanding> outstandings)
        {
            using (var fileData = new FileStream(fileName, FileMode.Create))
            {
                //_workbook.Write(fileData);
                _workbook = new XSSFWorkbook(fileData);
            }
            IFont font = (XSSFFont)_workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.FontName = "Courier New";

            ISheet sheet = _workbook.GetSheet("Report");

            // Defining a border
            var borderedCellStyle = BorderedCellStyle(font);

            var numberCellStyle = NumberCellStyle(font);
            //  numberCellStyle.DataFormat = numberDataFormat.GetFormat("");

            var dateCellStyle = DateCellStyle(font);

            //Creat The Headers of the excel
            //CreateHeaderRow(sheet, borderedCellStyle);
            CreateHeader(sheet, borderedCellStyle);
            int rowIndex = 2;
            foreach (var outstanding in outstandings)
            {
                IRow currentRow = sheet.CreateRow(rowIndex);

                CreateCell(currentRow, 0,outstanding.CustomerCode, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.CustomerName, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Consignee, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Product, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Reference, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.CustomerName, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Consignee, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Product, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.CustomerCode, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.CustomerName, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Consignee, borderedCellStyle);
                CreateCell(currentRow, 0, outstanding.Product, borderedCellStyle);


            }
        }

        private void CreateHeader(ISheet sheet, ICellStyle cellStyle)
        {
            IRow headerRow = sheet.CreateRow(0);
            CreateCell(headerRow, 0, "CUSTOMER", cellStyle);
            CreateCell(headerRow, 1, "CUSTOMER NAME", cellStyle);
            CreateCell(headerRow, 2, "CONSIGNEE", cellStyle);
            CreateCell(headerRow, 3, "PRODUCT", cellStyle);
            CreateCell(headerRow, 4, "REFERENCE", cellStyle);
            CreateCell(headerRow, 5, "INVOICE", cellStyle);
            CreateCell(headerRow, 6, "INVOICE DATE", cellStyle);
            CreateCell(headerRow, 7, "DUE DATE", cellStyle);
            CreateCell(headerRow, 8, "ARRER(DAYS)", cellStyle);
            CreateCell(headerRow, 9, "SLAB(DAYS)", cellStyle);
            CreateCell(headerRow, 10, "AMOUNT IN BDT", cellStyle);
        }
    }

    public interface IReporting
    {
        void CreateReport(string fileName, IEnumerable<Outstanding> outstandings);
    }
}