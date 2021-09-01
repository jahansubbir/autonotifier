using System;
using System.Collections.Generic;
using System.Linq;
using AutoNotifier.Models;

namespace AutoNotifier.BusinessLogic
{
    public class OutstandingDataManager
    {
        private readonly IExcelGateway _excelGateway;

        public OutstandingDataManager(IExcelGateway excelGateway)
        {
            _excelGateway = excelGateway;
        }

        public IEnumerable<Outstanding> GetOutstandingEnumerable(string fileName, string sheetName, string range)
        {
            var outstandings=  _excelGateway.GetOutstandings(fileName, sheetName, range);
            var outstandingList = outstandings.ToList();
                outstandingList.RemoveAll(a =>
                a.Product.IndexOf("vas", StringComparison.InvariantCultureIgnoreCase) >= 0 &&
                a.NetDueDate > DateTime.Today.AddDays(-10));
            return outstandingList;
        }

        public IEnumerable<IGrouping<string, Outstanding>> GetCustomerWiseOutstanding(
            IEnumerable<Outstanding> outstandings)
        {
            
            return outstandings.GroupBy(a => a.CustomerCode);
        }
        

    }
}
