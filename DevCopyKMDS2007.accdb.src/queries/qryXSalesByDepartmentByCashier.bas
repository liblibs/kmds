Operation =1
Option =0
Where ="(((InvoLine.TransType)<>'T' And (InvoLine.TransType)<>'X'))"
Having ="(((Cashier.ID)=[parCashierNumber]))"
Begin InputTables
    Name ="Cashier"
    Name ="InvoLine"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Expression ="SalesDepartment.ID"
    Expression ="SalesDepartment.Description"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Expression ="Cashier.ID"
    Expression ="Cashier.Name"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
    Flag =1
    LeftTable ="Cashier"
    RightTable ="InvoLine"
    Expression ="Cashier.ID = InvoLine.CashierID"
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
        Left =21
        Top =133
        Right =117
        Bottom =217
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =346
        Top =5
        Right =442
        Bottom =179
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =342
        Top =68
        Right =438
        Bottom =212
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =90
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
