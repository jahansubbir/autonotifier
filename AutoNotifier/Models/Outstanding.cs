using System;

namespace AutoNotifier.Models
{
    public class Outstanding
    {
        public string CustomerCode { get; set; }

        public string CustomerName { get; set; }
        public string Consignee { get; set; }
        public string Assignment { get; set; }
        public string DocumentNo { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DueDate { get; set; }
        public decimal Amount { get; set; }
        public string PayT { get; set; }
        public DateTime NetDueDate { get; set; }
        public string Slab { get; set; }
        public string Product { get; set; }
        public Double Arrear { get; set; }
        public string BillRefNumber { get; set; }
        public string DocumentType { get; set; }
    }
}
