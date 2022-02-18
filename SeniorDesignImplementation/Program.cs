using Gurobi;
using SeniorDesignImplementation;

var InventoryCapacity = 400000;
var alpha = 0.8;
var period = 24;

Console.WriteLine("Excel Get Data Started.");
var excelDataReceiver = new ExcelDataReceiver("BitirmeData24ay.xlsx");
var items = excelDataReceiver.GetItems();
var purchasableItems = items.Where(x => x.IsPurchasable).ToList();
var suppliers = excelDataReceiver.GetSuppliers();
excelDataReceiver.AssignItemSupplyProperties(suppliers, purchasableItems);
var itemDemands = excelDataReceiver.GetDemands(items);
var productionCapacities = excelDataReceiver.GetCapacities(items);
Console.WriteLine("Excel Get Data Completed.");

GRBEnv env = new GRBEnv();
GRBModel model = new GRBModel(env);

GRBVar[,,] PRist = new GRBVar[items.Count, suppliers.Count, period + 1];
GRBVar[,,] Vist = new GRBVar[items.Count, suppliers.Count, period + 1];
for (int i = 0; i < items.Count; i++)
{
    for (int s = 0; s < suppliers.Count; s++)
    {
        PRist[i, s, 0] = model.AddVar(0, 0, 0, GRB.INTEGER, $"PR_{i + 1}_{s + 1}_{0}");
        Vist[i, s, 0] = model.AddVar(0, 0, 0, GRB.BINARY, $"V_{i + 1}_{s + 1}_{0}");
        for (var t = 1; t <= period; t++)
        {
            PRist[i, s, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"PR_{i + 1}_{s + 1}_{t}");
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
    Xit[i, 0] = model.AddVar(0, 0, 0, GRB.INTEGER, $"X_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        Xit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"X_{i + 1}_{t}");
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
    Iit[i, 0] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"I_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        Iit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"I_{i + 1}_{t}");
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
    BOit[i, 0] = model.AddVar(0, 0, 0, GRB.INTEGER, $"BO_{i + 1}_{0}");
    Zit[i, 0] = model.AddVar(0, 0, 0, GRB.INTEGER, $"Z_{i + 1}_{0}");
    for (int t = 1; t <= period; t++)
    {
        BOit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"BO_{i + 1}_{t}");
        Zit[i, t] = model.AddVar(0, GRB.INFINITY, 0, GRB.INTEGER, $"Z_{i + 1}_{t}");
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
    //Z1t + I1t + X3t + X5t = I1t-1 + sum of s(PR1st) + X1t for all t
    GRBLinExpr lhs1 = new GRBLinExpr();
    lhs1.AddTerm(1, Zit[0, t]);
    lhs1.AddTerm(1, Iit[0, t]);
    lhs1.AddTerm(1, Xit[2, t]);
    lhs1.AddTerm(1, Xit[4, t]);
    GRBLinExpr rhs1 = new GRBLinExpr();
    rhs1.AddTerm(1, Iit[0, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs1.AddTerm(1, PRist[0, s, t]);
    }
    rhs1.AddTerm(1, Xit[0, t]);
    model.AddConstr(lhs1, GRB.EQUAL, rhs1, $"DependedOfItem1{1}_{t}");

    //Z2t + I2t + X4t + X5t = I2t-1 + sum of s(PR2st) + X2t for all t
    GRBLinExpr lhs2 = new GRBLinExpr();
    lhs2.AddTerm(1, Zit[1, t]);
    lhs2.AddTerm(1, Iit[1, t]);
    lhs2.AddTerm(1, Xit[3, t]);
    lhs2.AddTerm(1, Xit[4, t]);
    GRBLinExpr rhs2 = new GRBLinExpr();
    rhs2.AddTerm(1, Iit[1, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs2.AddTerm(1, PRist[1, s, t]);
    }
    rhs2.AddTerm(1, Xit[1, t]);
    model.AddConstr(lhs2, GRB.EQUAL, rhs2, $"DependedOfItem2{2}_{t}");

    //Z3t + I3t + X8t = I3t-1 + sum of s(PR3st) + X3t for all t
    GRBLinExpr lhs3 = new GRBLinExpr();
    lhs3.AddTerm(1, Zit[2, t]);
    lhs3.AddTerm(1, Iit[2, t]);
    lhs3.AddTerm(1, Xit[7, t]);
    GRBLinExpr rhs3 = new GRBLinExpr();
    rhs3.AddTerm(1, Iit[2, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs3.AddTerm(1, PRist[2, s, t]);
    }
    rhs3.AddTerm(1, Xit[2, t]);
    model.AddConstr(lhs3, GRB.EQUAL, rhs3, $"DependedOfItem3{3}_{t}");

    //Z4t + I4t + X7t + X8t = I4t-1 + sum of s(PR4st) + X4t for all t
    GRBLinExpr lhs4 = new GRBLinExpr();
    lhs4.AddTerm(1, Zit[3, t]);
    lhs4.AddTerm(1, Iit[3, t]);
    lhs4.AddTerm(1, Xit[6, t]);
    lhs4.AddTerm(1, Xit[7, t]);
    GRBLinExpr rhs4 = new GRBLinExpr();
    rhs4.AddTerm(1, Iit[3, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs4.AddTerm(1, PRist[3, s, t]);
    }
    rhs4.AddTerm(1, Xit[3, t]);
    model.AddConstr(lhs4, GRB.EQUAL, rhs4, $"DependedOfItem4{4}_{t}");

    //Z5t + I5t + X9t = I5t-1 + sum of s(PR5st) + X5t for all t
    GRBLinExpr lhs5 = new GRBLinExpr();
    lhs5.AddTerm(1, Zit[4, t]);
    lhs5.AddTerm(1, Iit[4, t]);
    lhs5.AddTerm(1, Xit[8, t]);
    GRBLinExpr rhs5 = new GRBLinExpr();
    rhs5.AddTerm(1, Iit[4, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs5.AddTerm(1, PRist[4, s, t]);
    }
    rhs5.AddTerm(1, Xit[4, t]);
    model.AddConstr(lhs5, GRB.EQUAL, rhs5, $"DependedOfItem5{5}_{t}");

    //Z6t + I6t + X10t = I6t-1 + sum of s(PR6st) + X6t for all t
    GRBLinExpr lhs6 = new GRBLinExpr();
    lhs6.AddTerm(1, Zit[5, t]);
    lhs6.AddTerm(1, Iit[5, t]);
    lhs6.AddTerm(1, Xit[9, t]);
    GRBLinExpr rhs6 = new GRBLinExpr();
    rhs6.AddTerm(1, Iit[5, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs6.AddTerm(1, PRist[5, s, t]);
    }
    rhs6.AddTerm(1, Xit[5, t]);
    model.AddConstr(lhs6, GRB.EQUAL, rhs6, $"DependedOfItem6{6}{t}");

    //Z7t + I7t + X9t + X10t = I7t-1 + sum of s(PR7st) + X7t for all t
    GRBLinExpr lhs7 = new GRBLinExpr();
    lhs7.AddTerm(1, Zit[6, t]);
    lhs7.AddTerm(1, Iit[6, t]);
    lhs7.AddTerm(1, Xit[8, t]);
    lhs7.AddTerm(1, Xit[9, t]);
    GRBLinExpr rhs7 = new GRBLinExpr();
    rhs7.AddTerm(1, Iit[6, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs7.AddTerm(1, PRist[6, s, t]);
    }
    rhs7.AddTerm(1, Xit[6, t]);
    model.AddConstr(lhs7, GRB.EQUAL, rhs7, $"DependedOfItem7{7}{t}");

    //Z8t + I8t + alpha[8,9]*X9t + X10t = I8t-1 + sum of s(PR8st)  + X8t for all t
    GRBLinExpr lhs8 = new GRBLinExpr();
    lhs8.AddTerm(1, Zit[7, t]);
    lhs8.AddTerm(1, Iit[7, t]);
    lhs8.AddTerm(1, Xit[8, t]);
    lhs8.AddTerm(1, Xit[9, t]);
    GRBLinExpr rhs8 = new GRBLinExpr();
    rhs8.AddTerm(1, Iit[7, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs8.AddTerm(1, PRist[7, s, t]);
    }
    rhs8.AddTerm(1, Xit[7, t]);
    model.AddConstr(lhs8, GRB.EQUAL, rhs8, $"DependedOfItem8{8}{t}");

    //Z9t + I9t = I9t-1 + sum of s(PR9st) + X9t
    GRBLinExpr lhs9 = new GRBLinExpr();
    lhs9.AddTerm(1, Zit[8, t]);
    lhs9.AddTerm(1, Iit[8, t]);
    GRBLinExpr rhs9 = new GRBLinExpr();
    rhs9.AddTerm(1, Iit[8, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs9.AddTerm(1, PRist[8, s, t]);
    }
    rhs9.AddTerm(1, Xit[8, t]);
    model.AddConstr(lhs9, GRB.EQUAL, rhs9, $"DependedOfItem9{9}{t}");

    //Z10t + I10t = I10t-1 + sum of s(PR10st) + X10t for all t
    GRBLinExpr lhs10 = new GRBLinExpr();
    lhs10.AddTerm(1, Zit[9, t]);
    lhs10.AddTerm(1, Iit[9, t]);
    GRBLinExpr rhs10 = new GRBLinExpr();
    rhs10.AddTerm(1, Iit[9, t - 1]);
    for (int s = 0; s < suppliers.Count; s++)
    {
        rhs10.AddTerm(1, PRist[9, s, t]);
    }
    rhs10.AddTerm(1, Xit[9, t]);
    model.AddConstr(lhs10, GRB.EQUAL, rhs10, $"DependedOfItem10{10}{t}");

}

//for (int i = 0; i < items.Count; i++)
//{
//    for(int t = 1; t <= 12; t++)
//    {
//        GRBLinExpr lhs = new GRBLinExpr();
//        lhs.AddTerm(1, Xit[i, t]);
//        lhs.AddTerm(1, Iit[i, t - 1]);
//        lhs.AddTerm(-1, Zit[i, t]);
//        for(int s = 0; s < suppliers.Count; s++)
//        {
//            lhs.AddTerm(1, PRist[i, s, t]);
//        }

//        GRBLinExpr rhs = new GRBLinExpr();
//        rhs.AddTerm(1, Iit[i, t]);
//    }
//}


//Objective Function
GRBLinExpr objectiveFunction = new GRBLinExpr();
for(int i = 0; i < items.Count; i++)
{
    for(int t = 1; t <= period; t++)
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
    double [,] array = new double[vars.GetLength(0), vars.GetLength(1)];
    for (int i = 0; i < vars.GetLength(0); i++)
    {
        for (int j = 1; j < vars.GetLength(1); j++)
        {
            array[i,j] = (int)vars[i, j].Get(GRB.DoubleAttr.X);
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