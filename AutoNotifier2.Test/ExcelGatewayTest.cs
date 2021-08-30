using System;
using System.Linq;
using Autonotifier2.DataAccess;
using Autonotifier2.Utilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AutoNotifier2.Test
{
    [TestClass]
    public class ExcelGatewayTest
    {
        public static string fileName =
            @"C:\Users\msj046\source\repos\AutoNotifier\AutoNotifier\bin\Debug\Resources\BD Weekly Outstanding Report -Week  32.xlsb";

        [TestMethod]
        public void GetOutstanding_TakesFile_ReturnList()
        {
            //arrange
          
        ExcelGateway excelGateway=new ExcelGateway(new JsonColumnHeaderReader());
            //act
            var data=excelGateway.GetOutstandings(fileName, "DATA", null);
            //assert
            Assert.IsTrue(data.Count()>1,"Exception happens");

            
        }
    }
}
