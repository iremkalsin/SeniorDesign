using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorDesignImplementation.Objects
{
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public double UnitaryHoldingCost { get; set; }
        public double UnitaryProducingCost { get; set; }
        public double SetupCost { get; set; }
        public List<int> PrerequisiteItems { get; set; }
        public Dictionary<int, double> FixedCostsOfItem { get; set; }
        public Dictionary<int, double> ProductionCostsOfItem { get; set; }
        public Dictionary<int, double> HoldingCostsOfItem { get; set; }
        public Dictionary<int, double> BackorderCostsOfItem { get; set; }
        public double UnitaryBackOrderCost { get; set; }
        public bool IsPurchasable { get; set; }

    }
}
