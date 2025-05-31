using System;
using System.Collections.Generic;

namespace StoreManagement.Models
{
    public class Order
    {
        public string CustomerName { get; set; }
        public string ShippingAddress { get; set; }
        public DateTime OrderDate { get; set; } = DateTime.Now;
        public List<OrderDetail> Items { get; set; } = new List<OrderDetail>();

        public decimal TotalAmount => Items.Sum(i => i.UnitPrice * i.Quantity);
    }
}
