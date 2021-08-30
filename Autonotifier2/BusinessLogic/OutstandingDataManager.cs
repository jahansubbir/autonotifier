using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autonotifier2.Models;

namespace Autonotifier2.BusinessLogic
{
    public class OutstandingDataManager
    {
        private readonly IExcelGateway _excelGateway;

        public OutstandingDataManager(IExcelGateway excelGateway)
        {
            _excelGateway = excelGateway;
        }

        public IEnumerable<IGrouping<string, Outstanding>> GetCustomerWiseOutstanding(string fileName, string sheetName, string range)
        {
            var outstandings = _excelGateway.GetOutstandings(fileName, sheetName, range);

            //filter out the DD type documents
            /*var documentTypeROustandings = allOutstandings.Where(a =>
                !a.DocumentType.StartsWith("d", StringComparison.InvariantCultureIgnoreCase));*/

            var groupByCustomer = outstandings.GroupBy(a => a.CustomerCode);
            return groupByCustomer;
        }

    }
}
