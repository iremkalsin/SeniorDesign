using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorDesignImplementation.Objects
{
    public class Supplier
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public Dictionary<Item, ItemSupplyProperties> Items { get; set; } = new Dictionary<Item, ItemSupplyProperties>();

    }
}
