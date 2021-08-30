using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autonotifier2.Models
{
   public interface IExcelGateway
    {
        IEnumerable<Outstanding> GetOutstandings(string fileName, string sheetName, string range);
    }
}
