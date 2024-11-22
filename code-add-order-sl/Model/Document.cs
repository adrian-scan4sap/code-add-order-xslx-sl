using System;

namespace code_add_order_sl.Model
{
    public class Document
    {
        public string CardCode { get; set; }
        public DateTime DocDueDate { get; set; }
        public int LineNum { get; set; }
        public string ItemCode { get; set; }
        public decimal Quantity { get; set; }
    }
}
