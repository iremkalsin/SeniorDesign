using Gurobi;
using SeniorDesignImplementation;

var InventoryCapacity = 400000;
var alpha = 0.8;
var period = 36;

Console.WriteLine("Excel Get Data Started.");
var excelDataReceiver = new ExcelDataReceiver("BitirmeData_36_40_Copy.xlsx");
var items = excelDataReceiver.GetItems();
var purchasableItems = items.Where(x => x.IsPurchasable).ToList();
var suppliers = excelDataReceiver.GetSuppliers();
excelDataReceiver.AssignItemSupplyProperties(suppliers, purchasableItems);
var itemDemands = excelDataReceiver.GetDemands(items);
var productionCapacities = excelDataReceiver.GetCapacities(items);
var coeffs = excelDataReceiver.TakeCoefficients();
Console.WriteLine("Excel Get Data Completed.");

GRBEnv env = new GRBEnv();
GRBModel model = new GRBModel(env);

GRBVar[,,] PRist = new GRBVar[items.Count, suppliers.Count, period + 1];
GRBVar[,,] Vist = new GRBVar[items.Count, suppliers.Count, period + 1];
for (int i = 0; i < items.Count; i++)
{
    for (int s = 0; s < suppliers.Count; s++)
    {
        PRist[i, s, 0] = model.AddVar(0, 0, 0, GRB.CONTINUOUS, $"PR_{i + 1}_{s + 1}_{0}");
        Vist[i, s, 0] = model.AddVar(0, 0, 0, GRB.BINARY, $"V_{i + 1}_{s + 1}_{0}");
        for (var t = 1; t <= period; t++)
        {
            PRist[i, s, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"PR_{i + 1}_{s + 1}_{t}");
            Vist[i, s, t] = model.AddVar(0, 1, 0, GRB.BINARY, $"V_{i + 1}_{s + 1}_{t}");
        }
    }
}

for (int i = 0; i < items.Count; i++) 
{
    for (int s = 0; s < suppliers.Count; s++)
    {
        for (var t = 1; t <= period; t++)
        {
            //PRist <= MXis * Vist for all i, s, t
            GRBLinExpr lhs = new GRBLinExpr();
            lhs.AddTerm(1, PRist[i, s, t]);
            GRBLinExpr rhs = new GRBLinExpr();
            rhs.AddTerm(suppliers.ElementAt(s).Items.ElementAt(i).Value.MaximumPurchaseAmount, Vist[i, s, t]);
            model.AddConstr(lhs, GRB.LESS_EQUAL, rhs, $"MaxPurchaseConst{i+1}_{s+1}_{t}");
            //MNis - PRist <= MNis*1 - MNis*Vist for all i, s, t
            GRBLinExpr lhs2 = new GRBLinExpr();
            lhs2.AddTerm(-1, PRist[i, s, t]);
            lhs2.AddConstant(suppliers.ElementAt(s).Items.ElementAt(i).Value.MinimumPurchaseAmount);
            GRBLinExpr rhs2 = new GRBLinExpr();
            rhs2.AddTerm(-suppliers.ElementAt(s).Items.ElementAt(i).Value.MinimumPurchaseAmount, Vist[i, s, t]);
            rhs2.AddConstant(suppliers.ElementAt(s).Items.ElementAt(i).Value.MinimumPurchaseAmount);
            model.AddConstr(lhs2, GRB.LESS_EQUAL, rhs2, $"MinORNonePurchaseConst{i+1}_{s+1}_{t}");
        }
    }
}

GRBVar[,] Xit = new GRBVar[items.Count, period + 1];
GRBVar[,] Yit = new GRBVar[items.Count, period + 1];
for (int i = 0; i < items.Count; i++)
{
    Yit[i, 0] = model.AddVar(0, 0, 0, GRB.BINARY, $"Y_{i + 1}_{0}");
    Xit[i, 0] = model.AddVar(0, 0, 0, GRB.CONTINUOUS, $"X_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        Xit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"X_{i + 1}_{t}");
        Yit[i, t] = model.AddVar(0, 1, 0, GRB.BINARY, $"Y_{i + 1}_{t}");
    }
}

for (int i = 0; i < items.Count; i++)
{
   
    for (int t = 1; t <= period; t++)
    {
        //Xit <= Cit*Yit for all i, t
        GRBLinExpr lhs2 = new GRBLinExpr();
        lhs2.AddTerm(1, Xit[i, t]);
        GRBLinExpr rhs2 = new GRBLinExpr();
        rhs2.AddTerm(productionCapacities.ElementAt(i).Value.ElementAt(t - 1).Amount, Yit[i, t]);
        model.AddConstr(lhs2, GRB.LESS_EQUAL, rhs2, $"ProduceORNotConst{i+1}_{t}");
    }
}

GRBVar[,] Iit = new GRBVar[items.Count, period + 1];
for (int i = 0; i < items.Count; i++)
{
    Iit[i, 0] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"I_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        Iit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"I_{i + 1}_{t}");
    }
}

for (int t = 1; t <= period; t++)
{
    GRBLinExpr lhs = new GRBLinExpr();
    for (int i = 0; i < items.Count; i++)
    {
        //sum of i(Iit)  <=  IC for all t
        lhs.AddTerm(1, Iit[i, t]);
    }
    model.AddConstr(lhs, GRB.LESS_EQUAL, InventoryCapacity, $"InventoryCapacityConst{t}");
}

GRBVar[,] BOit = new GRBVar[items.Count, period + 1];
GRBVar[,] Zit = new GRBVar[items.Count, period + 1];
for (int i = 0; i < items.Count; i++)
{
    BOit[i, 0] = model.AddVar(0, 0, 0, GRB.CONTINUOUS, $"BO_{i + 1}_{0}");
    Zit[i, 0] = model.AddVar(0, 0, 0, GRB.CONTINUOUS, $"Z_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        BOit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"BO_{i + 1}_{t}");
        Zit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.CONTINUOUS, $"Z_{i + 1}_{t}");
    }
}

for (int i = 0; i < items.Count; i++)
{
    for (int t = 1; t <= period; t++)
    {
        //BOit = BOi,t-1 + Di,t - Zi,t for all i, t
        GRBLinExpr lhs = new GRBLinExpr();
        lhs.AddTerm(1, BOit[i, t]);
        GRBLinExpr rhs = new GRBLinExpr();
        rhs.AddTerm(-1, Zit[i, t]);
        rhs.AddTerm(1, BOit[i, t - 1]);
        rhs.AddConstant(itemDemands.ElementAt(i).Value.ElementAt(t-1).Amount);
        model.AddConstr(lhs, GRB.EQUAL, rhs, $"BackOrderBalanceConst{i+1}_{t}");



        //demandin bir kısmı karşılanmalı
        //Zit >= ɑ*Dit
        GRBLinExpr lhs3 = new GRBLinExpr();
        lhs3.AddTerm(1, Zit[i, t]);
        GRBLinExpr rhs3 = new GRBLinExpr();
        rhs3.AddConstant(alpha * itemDemands.ElementAt(i).Value.ElementAt(t - 1).Amount);
        model.AddConstr(lhs3, GRB.GREATER_EQUAL, rhs3, $"SendSmallerThanDemandConst{i + 1}_{t}");
    }
}

for(int i = 0; i < items.Count; i++)
{
    //Ii0 = 0  Ii12 = 0  BOi12 = 0, BOi0 = 0 for all I
    model.AddConstr(Iit[i, 0], GRB.EQUAL, 0, $"InitialInvZero{i+1}");
    //model.AddConstr(Iit[i, period], GRB.EQUAL, 0, $"FinalInvZero{i+1}");
    model.AddConstr(BOit[i, 0], GRB.EQUAL, 0, $"InitialBOZero{i+1}");
    model.AddConstr(BOit[i, period], GRB.EQUAL, 0, $"FinalBOZero{i+1}");
}

for (int t = 1; t <= period; t++)
{

    for (int i = 0; i < items.Count; i++)
    {
        GRBLinExpr lhs = new GRBLinExpr();
        lhs.AddTerm(1, Zit[i, t]);
        lhs.AddTerm(1, Iit[i, t]);
        for (int j = 0; j < items.Count; j++)
        {
            for (int y = 0; y < items[j].PrerequisiteItems.Count; y++)
            {
                if (items[j].PrerequisiteItems[y] == items[i].Index)
                {
                    lhs.AddTerm(coeffs[i, j], Xit[items[j].Index - 1, t]);
                }
            }
        }
        GRBLinExpr rhs = new GRBLinExpr();
        rhs.AddTerm(1, Iit[i, t - 1]);
        for (int s = 0; s < suppliers.Count; s++)
        {
            rhs.AddTerm(1, PRist[i, s, t]);
        }
        rhs.AddTerm(1, Xit[i, t]);
        model.AddConstr(lhs, GRB.EQUAL, rhs, $"DependedOfItems{i}_{t}");
    }
}

    //Objective Function

    GRBLinExpr objectiveFunction = new GRBLinExpr();
    for (int i = 0; i < items.Count; i++)
    {
        for (int t = 1; t <= period; t++)
        {
            //min sum of t, i[(pit * xit)  + (hit * Iit) + (qit * Yit)] + sum of t, s, i(PRist * bis) + sum of i, t(BOit * BCit)
            objectiveFunction.AddTerm(items.ElementAt(i).ProductionCostsOfItem[t], Xit[i, t]);
            objectiveFunction.AddTerm(items.ElementAt(i).HoldingCostsOfItem[t], Iit[i, t]);
            objectiveFunction.AddTerm(items.ElementAt(i).FixedCostsOfItem[t], Yit[i, t]);
            objectiveFunction.AddTerm(items.ElementAt(i).BackorderCostsOfItem[t], BOit[i, t]);
            for (int s = 0; s < suppliers.Count; s++)
            {
                objectiveFunction.AddTerm(suppliers.ElementAt(s).Items.ElementAt(i).Value.UnitaryPurchasingCost, PRist[i, s, t]);
            }
        }
    }


    model.SetObjective(objectiveFunction, GRB.MINIMIZE);

    // Optimize
    model.Optimize();
    int status = model.Status;

    for (int i = 0; i < items.Count; i++)
    {
        for (int t = 1; t <= period; t++)
        {
            int I = (int)Iit[i, t].Get(GRB.DoubleAttr.X);
            int Z = (int)Zit[i, t].Get(GRB.DoubleAttr.X);
            int BO = (int)BOit[i, t].Get(GRB.DoubleAttr.X);
            int X = (int)Xit[i, t].Get(GRB.DoubleAttr.X);
            int Y = (int)Yit[i, t].Get(GRB.DoubleAttr.X);
            ExcelDataWriter writer = new ExcelDataWriter();
            writer.WriteToSheet("I", ConvertGRBVarArrayToDoubleArray(Iit));
            writer.WriteToSheet("Zit", ConvertGRBVarArrayToDoubleArray(Zit));
            writer.WriteToSheet("BOit", ConvertGRBVarArrayToDoubleArray(BOit));
            writer.WriteToSheet("Xit", ConvertGRBVarArrayToDoubleArray(Xit));
            writer.WriteToSheet("Yit", ConvertGRBVarArrayToDoubleArray(Yit));
            for (int s = 0; s < suppliers.Count; s++)
            {
                int PR = (int)PRist[i, s, t].Get(GRB.DoubleAttr.X);
                int V = (int)Vist[i, s, t].Get(GRB.DoubleAttr.X);
                writer.WriteToSheet($"Pi{s}t", ConvertGRBVarArrayToDoubleArray(ConvertThreeDArrayToTwoDArray(PRist, s)));
            }
            writer.Save();
        }
    }
    Console.WriteLine("Excel Write Completed.");

    Console.WriteLine();

    if (status == GRB.Status.UNBOUNDED)
    {
        Console.WriteLine("The model cannot be solved " + " because it is unbounded ");
        return;
    }
    if (status == GRB.Status.OPTIMAL)
    {
        Console.WriteLine("The optimal objective is " + model.ObjVal);
        return;
    }
    if ((status != GRB.Status.INF_OR_UNBD) && (status != GRB.Status.INFEASIBLE))
    {
        Console.WriteLine(" Optimization was stopped with status " + status);
        return;
    }

    double[,] ConvertGRBVarArrayToDoubleArray(GRBVar[,] vars)
    {
        double[,] array = new double[vars.GetLength(0), vars.GetLength(1)];
        for (int i = 0; i < vars.GetLength(0); i++)
        {
            for (int j = 1; j < vars.GetLength(1); j++)
            {
                array[i, j] = (int)vars[i, j].Get(GRB.DoubleAttr.X);
            }
        }
        return array;
    }

    GRBVar[,] ConvertThreeDArrayToTwoDArray(GRBVar[,,] vars, int staticIndex)
    {
        GRBVar[,] twoDArray = new GRBVar[vars.GetLength(0), vars.GetLength(2)];
        for (int i = 0; i < vars.GetLength(0); i++)
        {
            for (int t = 1; t < vars.GetLength(2); t++)
            {
                twoDArray[i, t] = vars[i, staticIndex, t];
            }
        }
        return twoDArray;
    }
