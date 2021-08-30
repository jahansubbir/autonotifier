using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autonotifier2.BusinessLogic;
using Autonotifier2.DataAccess;
using Autonotifier2.Utilities;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AutoNotifier2.Test
{
    [TestClass]
    public class OutstandingDataManagerTest
    {
        [TestMethod]
        public void GetCustomerWiseOutstanding_TakeFileName_ReturnIGrouping()
        {
            OutstandingDataManager outstandingDataManager=new OutstandingDataManager(
                new ExcelGateway(
                    new JsonColumnHeaderReader()));
            //act
            var data = outstandingDataManager.GetCustomerWiseOutstanding(ExcelGatewayTest.fileName, "DATA", null);

            //assert
           Assert.IsTrue(data.Any());
        Assert.IsTrue(data.All(a=>!a.ToList().Any(o => o.DocumentType.StartsWith("d",StringComparison.InvariantCultureIgnoreCase))));
        }
    }
}
