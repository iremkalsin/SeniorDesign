using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SeniorDesignImplementation.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace SeniorDesignImplementation
{
    public class ExcelDataReceiver
    {
        static int itemCount;

        public XSSFWorkbook workbook;
        public ExcelDataReceiver(string filePath)
        {
            using (FileStream file = new FileStream(@filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }
        }

        public ISheet GetWorksheet(string worksheetName)
        {
            return workbook.GetSheet(worksheetName);
        }

        public Dictionary<Item, List<Demand>> GetDemands(List<Item> items)
        {
            Dictionary<Item, List<Demand>> demands = new Dictionary<Item, List<Demand>>();
            var excelDemandsWorksheet = GetWorksheet("Demands");
            for (int rowNumber = 1; rowNumber <= excelDemandsWorksheet.LastRowNum; rowNumber++)
            {
                var periodicDemands = new List<Demand>();
                if (excelDemandsWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    var currentRow = excelDemandsWorksheet.GetRow(rowNumber);
                    for(int i = 1; i < currentRow.LastCellNum; i++)
                    {
                        periodicDemands.Add(new Demand()
                        {
                            Period = i,
                            Amount = currentRow.GetCell(i).NumericCellValue
                        });
                    }
                    demands.Add(items[rowNumber - 1], periodicDemands);
                    Console.WriteLine(currentRow);
                }
            }
            return demands;
        }

        internal Dictionary<Item, List<Capacity>> GetCapacities(List<Item> items)
        {
            Dictionary<Item, List<Capacity>> productionCapacities = new Dictionary<Item, List<Capacity>>();
            var excelDemandsWorksheet = GetWorksheet("ProductionCapacity");
            for (int rowNumber = 1; rowNumber <= excelDemandsWorksheet.LastRowNum; rowNumber++)
            {
                var periodicProductionCapacities = new List<Capacity>();
                if (excelDemandsWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    var currentRow = excelDemandsWorksheet.GetRow(rowNumber);
                    for (int i = 1; i < currentRow.LastCellNum; i++)
                    {
                        periodicProductionCapacities.Add(new Capacity()
                        {
                            Period = i,
                            Amount = currentRow.GetCell(i).NumericCellValue
                        });
                    }
                    productionCapacities.Add(items[rowNumber - 1], periodicProductionCapacities);
                    Console.WriteLine(currentRow);
                }
            }
            return productionCapacities;
        }



        //dasdasfa


        public List<Item> GetItems()
        {
            List<Item> items = new List<Item>();
            var excelItemsWorksheet = GetWorksheet("Items");
            var excelFixedCostsWorksheet = GetWorksheet("FixedCosts");
            var excelProductionCostsWorksheet = GetWorksheet("ProductionCosts");
            var excelHoldingCostsWorksheet = GetWorksheet("HoldingCosts");
            var excelBackorderCostsWorksheet = GetWorksheet("BackorderCosts");
            var excelPreReqWorksheet = GetWorksheet("PrerequisiteItems");
            for (int rowNumber = 1; rowNumber <= excelItemsWorksheet.LastRowNum; rowNumber++)
            {
                if (excelItemsWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    var currentRow = excelItemsWorksheet.GetRow(rowNumber);

                    Dictionary<int, double> fixedCosts = new Dictionary<int, double>();
                    Dictionary<int, double> productionCosts = new Dictionary<int, double>();
                    Dictionary<int, double> holdingCosts = new Dictionary<int, double>();
                    Dictionary<int, double> backorderCosts = new Dictionary<int, double>();

                    for (int c = 1; c < excelFixedCostsWorksheet.GetRow(rowNumber).LastCellNum; c++)
                    {
                        fixedCosts.Add((int)excelFixedCostsWorksheet.GetRow(0).GetCell(c).NumericCellValue, excelFixedCostsWorksheet.GetRow(rowNumber).GetCell(c).NumericCellValue);
                        productionCosts.Add((int)excelProductionCostsWorksheet.GetRow(0).GetCell(c).NumericCellValue, excelProductionCostsWorksheet.GetRow(rowNumber).GetCell(c).NumericCellValue);
                        holdingCosts.Add((int)excelHoldingCostsWorksheet.GetRow(0).GetCell(c).NumericCellValue, excelHoldingCostsWorksheet.GetRow(rowNumber).GetCell(c).NumericCellValue);
                        backorderCosts.Add((int)excelBackorderCostsWorksheet.GetRow(0).GetCell(c).NumericCellValue, excelBackorderCostsWorksheet.GetRow(rowNumber).GetCell(c).NumericCellValue);
                    }
                    var cellNum = 2;
                    List<int> preReqItems = new List<int>();
                    while(excelPreReqWorksheet.GetRow(rowNumber).GetCell(cellNum) != null)
                    {
                        preReqItems.Add((int)excelPreReqWorksheet.GetRow(rowNumber).GetCell(cellNum).NumericCellValue);
                        cellNum++;
                    }
                    
                    items.Add(new Item()
                    {  
                        Name = currentRow.GetCell(0).StringCellValue,
                        Index = (int)currentRow.GetCell(1).NumericCellValue,
                        FixedCostsOfItem = fixedCosts,
                        ProductionCostsOfItem = productionCosts,
                        HoldingCostsOfItem = holdingCosts,
                        BackorderCostsOfItem = backorderCosts,
                        PrerequisiteItems = preReqItems,
                        IsPurchasable = currentRow.GetCell(7).NumericCellValue == 1 ? true : false
                    });
                    Console.WriteLine(currentRow);
                }
            }
            itemCount = items.Count;
            return items;
        }
        public List<Supplier> GetSuppliers()
        {
            var suppliers = new List<Supplier>();
            var excelItemsWorksheet = GetWorksheet("Suppliers");
            for (int rowNumber = 1; rowNumber <= excelItemsWorksheet.LastRowNum; rowNumber++)
            {
                if (excelItemsWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    var currentRow = excelItemsWorksheet.GetRow(rowNumber);
                    suppliers.Add(new Supplier()
                    {
                        Index = rowNumber,
                        Name = currentRow.GetCell(0).StringCellValue
                    });
                    Console.WriteLine(currentRow);
                }
            }
            return suppliers;
        }

        public void AssignItemSupplyProperties(List<Supplier> suppliers, List<Item> items)
        {
            var excelPricesWorksheet = GetWorksheet("Prices");
            var excelMinimumAmountsWorksheet = GetWorksheet("MinimumAmounts");
            var excelMaximumAmountsWorksheet = GetWorksheet("MaximumAmounts");
            for (int rowNumber = 1; rowNumber <= excelPricesWorksheet.LastRowNum; rowNumber++)
            {
                if (excelPricesWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    var currentPriceRow = excelPricesWorksheet.GetRow(rowNumber);
                    var minimumAmountsRow = excelMinimumAmountsWorksheet.GetRow(rowNumber);
                    var maximumAmountsRow = excelMaximumAmountsWorksheet.GetRow(rowNumber);
                    var item = items[rowNumber - 1];
                    for(int i = 0; i < suppliers.Count; i++)
                    {
                        var itemSupplyProperty = new ItemSupplyProperties();
                        itemSupplyProperty.UnitaryPurchasingCost = currentPriceRow.GetCell(i + 1).NumericCellValue;
                        itemSupplyProperty.MinimumPurchaseAmount = minimumAmountsRow.GetCell(i + 1).NumericCellValue;
                        itemSupplyProperty.MaximumPurchaseAmount = maximumAmountsRow.GetCell(i + 1).NumericCellValue;
                        suppliers[i].Items.Add(item, itemSupplyProperty);
                    }
                    Console.WriteLine(currentPriceRow);
                }
            }
        }

        public double[,] TakeCoefficients()
        {
            double[,] coefficients = new double[itemCount, itemCount];
            var excelTransCoeffsWorksheet = GetWorksheet("TransformationCoeffs");
            for (int rowNumber = 1; rowNumber <= excelTransCoeffsWorksheet.LastRowNum; rowNumber++)
            {
                if (excelTransCoeffsWorksheet.GetRow(rowNumber) != null) //null is when the row only contains empty cells 
                {
                    for (int columnNumber = 1; columnNumber <= excelTransCoeffsWorksheet.LastRowNum; columnNumber++)
                    {
                        var currentRow = excelTransCoeffsWorksheet.GetRow(rowNumber);
                        coefficients[rowNumber - 1, columnNumber - 1] = currentRow.GetCell(columnNumber).NumericCellValue;
                    }
                }
            }
            return coefficients;
        }
    }
}
