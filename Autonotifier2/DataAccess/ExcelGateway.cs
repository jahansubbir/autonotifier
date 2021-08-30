using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autonotifier2.Models;
using Autonotifier2.Utilities;

namespace Autonotifier2.DataAccess
{
   public  class ExcelGateway : Gateway, IExcelGateway
   {
       private readonly JsonColumnHeaderReader _columnHeaderReader;//{ get; }

        public ExcelGateway(JsonColumnHeaderReader columnHeaderReader)
        {
            _columnHeaderReader = columnHeaderReader;
        }
        public IEnumerable<Outstanding> GetOutstandings(string fileName, string sheetName, string range)
        {
            List<Outstanding> outstandings = new List<Outstanding>();

            var columns = _columnHeaderReader.GetColumnDictionary();

            StringBuilder queryBuilder = new StringBuilder("SELECT TRIM(");
            foreach (var column in columns)
            {
                string columnName;
                if (column.Value.Contains("."))
                {
                    columnName = column.Value.Replace(".", "#");
                }
                else
                {
                    columnName = column.Value;
                }
                queryBuilder.Append("[");
                queryBuilder.Append(columnName);
                queryBuilder.Append("])");
                queryBuilder.Append(" as ");
                queryBuilder.Append(column.Key);
                queryBuilder.Append(",");

            }

            queryBuilder.Remove(queryBuilder.Length - 1, 1);
            
            queryBuilder.Append(" FROM ");
            queryBuilder.Append("[");
            queryBuilder.Append(sheetName);
            queryBuilder.Append("$");
            if (!string.IsNullOrEmpty(range))
            {
                queryBuilder.Append(range);
            }
            queryBuilder.Append("]");
            queryBuilder.Append(" WHERE [CoCd]<>'' AND");
            queryBuilder.Append("[");
            queryBuilder.Append(columns["DocumentType"]);
            queryBuilder.Append("]");
            queryBuilder.Append("<>'%D%'");
            Query = queryBuilder.ToString();

            Connect(fileName);
            Connection.Open();
            Command = new OleDbCommand(Query, Connection);
            Reader = Command.ExecuteReader();
            while (Reader.Read())
            {
                Outstanding outstanding = new Outstanding()
                {
                    Assignment = Reader["Assignment"]?.ToString(),
                    Consignee = Reader["Consignee"]?.ToString(),
                    CustomerCode = Reader["CustomerCode"]?.ToString(),
                    CustomerName = Reader["CustomerName"]?.ToString(),
                    DocumentType=Reader["DocumentType"]?.ToString(),
                    DocumentNo = Reader["DocumentNo"]?.ToString(),
                    PayT = Reader["PayT"]?.ToString(),
                    Product = Reader["Product"]?.ToString(),
                    Slab = Reader["Slab"]?.ToString(),
                };
                if (Double.TryParse(Reader["Arrear"].ToString(), out var arrear))
                {
                    outstanding.Arrear = arrear;
                }
                if (Decimal.TryParse(Reader["Amount"].ToString(), out var amount))
                {
                    outstanding.Amount = amount;
                }

                if (DateTime.TryParse(Reader["DocDate"].ToString(), out var docDate))
                {
                    outstanding.DocDate = docDate;

                }

                /*if (DateTime.TryParse(Reader["DueDate"].ToString(), out var dueDate))
                {
                    outstanding.DueDate = dueDate;

                }*/

                if (DateTime.TryParse(Reader["NetDueDate"].ToString(), out var netDueDate))
                {
                    outstanding.NetDueDate = netDueDate;
                }

                outstandings.Add(outstanding);

            }
            Reader.Close();
            Connection.Close();


            return outstandings;
        }
    }
}
