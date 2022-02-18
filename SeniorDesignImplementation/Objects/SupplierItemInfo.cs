using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorDesignImplementation.Objects
{
    public class SupplierItemInfo
    {
        public Supplier Supplier { get; set; }
        public Dictionary<Item, ItemSupplyProperties> ItemSupplyProperties { get; set; }
    }
}
