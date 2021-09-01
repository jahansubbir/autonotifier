using System.Collections.Generic;

namespace AutoNotifier.Models
{
   public interface IExcelGateway
    {
        IEnumerable<Outstanding> GetOutstandings(string fileName, string sheetName, string range);
    }
}
