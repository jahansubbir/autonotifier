using System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AutoNotifier.BusinessLogic
{
    public class NpoiExcel
    {
        protected IWorkbook _workbook;



        protected void CreateCell(IRow currentRow, int cellIndex, string value, ICellStyle style)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }
        protected void CreateCell(IRow currentRow, int cellIndex, double value, ICellStyle style)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }
        protected void CreateCell(IRow currentRow, int cellIndex, DateTime value, ICellStyle style)
        {
            ICell cell = currentRow.CreateCell(cellIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }
        /*public void InsertBLData(string fileName, List<CarrierBL> carrierBls)
        {
            using (var fileData = new FileStream(@"Resources\BLAUDIT.xlsx", FileMode.Open))
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


            int rowIndex = 2;
            // IRow dataRow = sheet.CreateRow(0);
            foreach (CarrierBL bl in carrierBls)
            {
                IRow currentRow = sheet.CreateRow(rowIndex);

                CreateCell(currentRow, 0, bl.BlNo, borderedCellStyle);
                CreateCell(currentRow, 2, bl.Shipper.Trim(), borderedCellStyle);
                CreateCell(currentRow, 4, bl.ShipperAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 6, bl.Consignee.Trim(), borderedCellStyle);
                CreateCell(currentRow, 8, bl.ConsigneeAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 10, bl.NotifyParty.Trim(), borderedCellStyle);
                CreateCell(currentRow, 12, bl.NotifyAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 14, bl.Vessel.Trim(), borderedCellStyle);
                CreateCell(currentRow, 16, bl.Voyage.Trim(), borderedCellStyle);
                CreateCell(currentRow, 18, bl.PortOfLoading.Trim(), borderedCellStyle);
                CreateCell(currentRow, 20, bl.PortOfDischarge.Trim(), borderedCellStyle);
                CreateCell(currentRow, 22, bl.PlaceOfReceipt.Trim(), borderedCellStyle);
                CreateCell(currentRow, 24, bl.PlaceOfDelivery.Trim(), borderedCellStyle);
                CreateCell(currentRow, 26, bl.Packages, borderedCellStyle);
                CreateCell(currentRow, 28, bl.Weight, numberCellStyle);
                CreateCell(currentRow, 30, bl.Cbm, numberCellStyle);
                CreateCell(currentRow, 32, bl.ContainerCount, numberCellStyle);
                CreateCell(currentRow, 34, bl.ShippedOnBoard, dateCellStyle);
                CreateCell(currentRow, 36, bl.IssueDate, dateCellStyle);
                CreateCell(currentRow, 38, bl.Description, dateCellStyle);

                rowIndex++;
            }
            int lastColumNum = sheet.GetRow(0).LastCellNum;
            for (int i = 0; i <= lastColumNum; i++)
            {
                sheet.AutoSizeColumn(i);
                GC.Collect();
            }

            if (!fileName.Contains(".xlsx"))
            {
                fileName = fileName + ".xlsx";
            }
            using (var fileData = new FileStream(fileName, FileMode.OpenOrCreate))
            {
                _workbook.Write(fileData);
                // fileData.Close();

                //_workbook = new XSSFWorkbook(fileData);
            }


        }*/

        protected void CreateHeaderRow(ISheet sheet, ICellStyle borderedCellStyle)
        {
            IRow headerRow = sheet.CreateRow(0);

            CreateCell(headerRow, 1, "BL", borderedCellStyle);
            CreateCell(headerRow, 3, "SHIPPER", borderedCellStyle);
            CreateCell(headerRow, 5, "SHIPPER ADDRESS", borderedCellStyle);
            CreateCell(headerRow, 7, "CONSIGNEE", borderedCellStyle);
            CreateCell(headerRow, 9, "CONSIGNEE ADDRESS", borderedCellStyle);
            CreateCell(headerRow, 11, "NOTIFY", borderedCellStyle);
            CreateCell(headerRow, 13, "NOTIFY ADDRESS", borderedCellStyle);
            CreateCell(headerRow, 15, "VESSEL", borderedCellStyle);
            CreateCell(headerRow, 17, "VOYAGE", borderedCellStyle);
            CreateCell(headerRow, 19, "PORT OF LOADING", borderedCellStyle);
            CreateCell(headerRow, 21, "PORT OF DISCHARGE", borderedCellStyle);
            CreateCell(headerRow, 23, "PLACE OF RECEIPT", borderedCellStyle);
            CreateCell(headerRow, 25, "PLACE OF DELIVERY", borderedCellStyle);
            CreateCell(headerRow, 27, "PACKAGE", borderedCellStyle);
            CreateCell(headerRow, 29, "WEIGHT", borderedCellStyle);
            CreateCell(headerRow, 31, "CBM", borderedCellStyle);
            CreateCell(headerRow, 33, "CONTAINER COUNT", borderedCellStyle);
            CreateCell(headerRow, 35, "ISSUE DATE", borderedCellStyle);
            CreateCell(headerRow, 37, "SHIPPED ON BOARD", borderedCellStyle);
        }

        protected ICellStyle BorderedCellStyle(IFont font)
        {
            ICellStyle borderedCellStyle = (XSSFCellStyle)_workbook.CreateCellStyle();
            borderedCellStyle.SetFont(font);
            borderedCellStyle.BorderLeft = BorderStyle.Medium;
            borderedCellStyle.BorderTop = BorderStyle.Medium;
            borderedCellStyle.BorderRight = BorderStyle.Medium;
            borderedCellStyle.BorderBottom = BorderStyle.Medium;
            borderedCellStyle.VerticalAlignment = VerticalAlignment.Center;
            return borderedCellStyle;
        }

        protected ICellStyle DateCellStyle(IFont font)
        {
            ICellStyle dateCellStyle = (XSSFCellStyle)_workbook.CreateCellStyle();
            IDataFormat dateFormat = (XSSFDataFormat)_workbook.CreateDataFormat();
            dateCellStyle.DataFormat = dateFormat.GetFormat("dd-MMM-yyyy");
            dateCellStyle.SetFont(font);
            dateCellStyle.BorderLeft = BorderStyle.Medium;
            dateCellStyle.BorderTop = BorderStyle.Medium;
            dateCellStyle.BorderRight = BorderStyle.Medium;
            dateCellStyle.BorderBottom = BorderStyle.Medium;
            dateCellStyle.VerticalAlignment = VerticalAlignment.Center;
            return dateCellStyle;
        }

        protected ICellStyle NumberCellStyle(IFont font)
        {
            ICellStyle numberCellStyle = (XSSFCellStyle)_workbook.CreateCellStyle();
            IDataFormat numberDataFormat = (XSSFDataFormat)_workbook.CreateDataFormat();
            numberCellStyle.SetFont(font);
            numberCellStyle.BorderLeft = BorderStyle.Medium;
            numberCellStyle.BorderTop = BorderStyle.Medium;
            numberCellStyle.BorderRight = BorderStyle.Medium;
            numberCellStyle.BorderBottom = BorderStyle.Medium;
            numberCellStyle.VerticalAlignment = VerticalAlignment.Center;
            return numberCellStyle;
        }

        /*public void InsertFcrData(string fileName, Dictionary<string, CarrierBL> fcrInfos)
        {
            if (!fileName.EndsWith(".xlsx"))
            {
                fileName += ".xlsx";
            }
            //_workbook = new XSSFWorkbook();
            using (var fileData = new FileStream(fileName, FileMode.Open))
            {
                _workbook = new XSSFWorkbook(fileData);//(fileData);
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
            //CreateHeaderRow(sheet,borderedCellStyle);


            int rowIndex = 2;
            // IRow dataRow = sheet.CreateRow(0);
            foreach (var blValue in fcrInfos)
            {
                var bl = blValue.Value;
                IRow currentRow = GetBlRow(sheet, blValue.Key);

                CreateCell(currentRow, 1, bl.BlNo, borderedCellStyle);
                CreateCell(currentRow, 3, bl.Shipper.Trim(), borderedCellStyle);
                CreateCell(currentRow, 5, bl.ShipperAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 7, bl.Consignee.Trim(), borderedCellStyle);
                CreateCell(currentRow, 9, bl.ConsigneeAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 11, bl.NotifyParty.Trim(), borderedCellStyle);
                CreateCell(currentRow, 13, bl.NotifyAddress.Trim(), borderedCellStyle);
                CreateCell(currentRow, 15, bl.Vessel.Trim(), borderedCellStyle);
                CreateCell(currentRow, 17, bl.Voyage.Trim(), borderedCellStyle);
                CreateCell(currentRow, 19, bl.PortOfLoading.Trim(), borderedCellStyle);
                CreateCell(currentRow, 21, bl.PortOfDischarge.Trim(), borderedCellStyle);
                CreateCell(currentRow, 23, bl.PlaceOfReceipt.Trim(), borderedCellStyle);
                CreateCell(currentRow, 25, bl.PlaceOfDelivery.Trim(), borderedCellStyle);
                CreateCell(currentRow, 27, bl.Packages, borderedCellStyle);
                CreateCell(currentRow, 29, bl.Weight, numberCellStyle);
                CreateCell(currentRow, 31, bl.Cbm, numberCellStyle);
                CreateCell(currentRow, 33, bl.ContainerCount, numberCellStyle);
                CreateCell(currentRow, 35, bl.ShippedOnBoard, dateCellStyle);
                CreateCell(currentRow, 37, bl.IssueDate, dateCellStyle);
                CreateCell(currentRow, 39, bl.Description, borderedCellStyle);
                //for (int i = 38; i < 51; i++)

                SetFormula(currentRow, rowIndex);


                rowIndex++;
            }
            int lastColumNum = sheet.GetRow(0).LastCellNum;

            for (int i = 0; i <= lastColumNum; i++)
            {
                sheet.AutoSizeColumn(i);
                GC.Collect();
            }
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(_workbook);

            if (!fileName.Contains(".xlsx"))
            {
                fileName = fileName + ".xlsx";
            }
            using (var fileData = new FileStream(fileName, FileMode.Open))
            {
                _workbook.Write(fileData);

                fileData.Close();
                //_workbook = new XSSFWorkbook(fileData);
            }

        }*/

        private IRow GetBlRow(ISheet sheet, string blValueKey)
        {
            IRow row = null;
            var lastRow = sheet.LastRowNum;
            for (int i = 0; i <= lastRow; i++)
            {
                var value = sheet.GetRow(i).GetCell(0).StringCellValue;
                if (value == blValueKey)
                {
                    row = sheet.GetRow(i);
                    break;
                }
            }

            return row;

        }

        protected static void SetFormula(IRow currentRow, int rowIndex)
        {
            rowIndex += 1;
            var cell = currentRow.CreateCell(40);
            cell.SetCellType(CellType.Formula);
            cell.SetCellFormula(String.Format("IF($C$" + rowIndex + "=$D$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell2 = currentRow.CreateCell(41);
            cell2.SetCellType(CellType.Formula);
            cell2.SetCellFormula(String.Format("IF($E$" + rowIndex + "=$F$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell3 = currentRow.CreateCell(42);
            cell3.SetCellType(CellType.Formula);
            cell3.SetCellFormula(String.Format("IF($G$" + rowIndex + "=$H$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell4 = currentRow.CreateCell(43);
            cell4.SetCellType(CellType.Formula);
            cell4.SetCellFormula(String.Format("IF($I$" + rowIndex + "=$J$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell5 = currentRow.CreateCell(44);
            cell5.SetCellType(CellType.Formula);
            cell5.SetCellFormula(String.Format("IF($K$" + rowIndex + "=$L$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell6 = currentRow.CreateCell(45);
            cell6.SetCellType(CellType.Formula);
            cell6.SetCellFormula(String.Format("IF($M$" + rowIndex + "=$N$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell7 = currentRow.CreateCell(46);
            cell7.SetCellType(CellType.Formula);
            cell7.SetCellFormula(String.Format("IF($O$" + rowIndex + "=$P$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell8 = currentRow.CreateCell(47);
            cell8.SetCellType(CellType.Formula);
            cell8.SetCellFormula(String.Format("IF($Q$" + rowIndex + "=$R$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell9 = currentRow.CreateCell(48);
            cell9.SetCellType(CellType.Formula);
            cell9.SetCellFormula(String.Format("IF($S$" + rowIndex + "=$T$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell10 = currentRow.CreateCell(49);
            cell10.SetCellType(CellType.Formula);
            cell10.SetCellFormula(String.Format("IF($U$" + rowIndex + "=$V$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell11 = currentRow.CreateCell(50);
            cell11.SetCellType(CellType.Formula);
            cell11.SetCellFormula(String.Format("IF($W$" + rowIndex + "=$X$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell12 = currentRow.CreateCell(51);
            cell12.SetCellType(CellType.Formula);
            cell12.SetCellFormula(String.Format("IF($Y$" + rowIndex + "=$Z$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell13 = currentRow.CreateCell(52);
            cell13.SetCellType(CellType.Formula);
            cell13.SetCellFormula(String.Format("IF($AA$" + rowIndex + "=$AB$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell14 = currentRow.CreateCell(53);
            cell14.SetCellType(CellType.Formula);
            cell14.SetCellFormula(String.Format("IF($AC$" + rowIndex + "=$AD$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell15 = currentRow.CreateCell(54);
            cell15.SetCellType(CellType.Formula);
            cell15.SetCellFormula(String.Format("IF($AE$" + rowIndex + "=$AF$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell16 = currentRow.CreateCell(55);
            cell16.SetCellType(CellType.Formula);
            cell16.SetCellFormula(String.Format("IF($AG$" + rowIndex + "=$AH$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell17 = currentRow.CreateCell(56);
            cell17.SetCellType(CellType.Formula);
            cell17.SetCellFormula(String.Format("IF($AI$" + rowIndex + "=$AJ$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell18 = currentRow.CreateCell(57);
            cell18.SetCellType(CellType.Formula);
            cell18.SetCellFormula(String.Format("IF($AK$" + rowIndex + "=$AL$" + rowIndex + ",\"YES\",\"NO\")"));
            var cell19 = currentRow.CreateCell(58);
            cell18.SetCellType(CellType.Formula);
            cell18.SetCellFormula(String.Format("IF($AM$" + rowIndex + "=$AN$" + rowIndex + ",\"YES\",\"NO\")"));
        }
    }
}
