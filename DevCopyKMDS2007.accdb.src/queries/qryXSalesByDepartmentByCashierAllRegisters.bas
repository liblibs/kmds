Operation =1
Option =0
Where ="(((InvoLineAllRegisters.TransType)<>'T' And (InvoLineAllRegisters.TransType)<>'X"
    "'))"
Having ="(((Cashier.ID)=[parCashierNumber]))"
Begin InputTables
    Name ="InvoLineAllRegisters"
    Name ="Item"
    Name ="SalesDepartment"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLineAllRegisters.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLineAllRegisters.Price)"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Item"
    Expression ="InvoLineAllRegisters.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="InvoLineAllRegisters"
    RightTable ="Cashier"
    Expression ="InvoLineAllRegisters.CashierID = Cashier.ID"
    Flag =1
End
Begin OrderBy
    Expression ="SalesDepartment.ID"
    Flag =0
    Expression ="Cashier.ID"
    Flag =0
End
Begin Groups
    Expression ="SalesDepartment.ID"
    GroupLevel =0
    Expression ="SalesDepartment.Description"
    GroupLevel =0
    Expression ="Cashier.ID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =245
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLineAllRegisters"
        Name =""
    End
    Begin
        Left =419
        Top =47
        Right =515
        Bottom =191
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =246
        Top =151
        Right =342
        Bottom =235
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =869
        Top =55
        Right =965
        Bottom =139
        Top =0
        Name ="Cashier"
        Name =""
    End
End
