Operation =1
Option =0
Where ="(((InvoLine5.TransType)<>'T'))"
Begin InputTables
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="InvoLine5"
    Name ="Item"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="InvoLine5.CashierID"
    Expression ="Cashier.Name"
    Alias ="Group"
    Expression ="SalesGroup.ID"
    Alias ="GroupDesc"
    Expression ="SalesGroup.Description"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Alias ="DepartmentDesc"
    Expression ="IIf(Involine5.Transtype='C',\"Club Sales\",SalesDepartment.Description)"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine5.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine5.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf([CashsysID]=1,[Quantity],0))"
    Alias ="D1"
    Expression ="Sum(IIf([CashsysID]=1,InvoLine5.Price,0))"
    Alias ="Q2"
    Expression ="Sum(IIf([CashsysID]=2,[Quantity],0))"
    Alias ="D2"
    Expression ="Sum(IIf([CashsysID]=2,InvoLine5.Price,0))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3,[Quantity],0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,InvoLine5.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4,[Quantity],0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,InvoLine5.Price,0))"
    Expression ="InvoLine5.TransType"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Cashier"
    Expression ="InvoLine5.CashierID = Cashier.ID"
    Flag =1
End
Begin OrderBy
    Expression ="InvoLine5.CashierID"
    Flag =0
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
End
Begin Groups
    Expression ="InvoLine5.CashierID"
    GroupLevel =0
    Expression ="Cashier.Name"
    GroupLevel =0
    Expression ="SalesGroup.ID"
    GroupLevel =0
    Expression ="SalesGroup.Description"
    GroupLevel =0
    Expression ="SalesDepartment.ID"
    GroupLevel =0
    Expression ="IIf(Involine5.Transtype='C',\"Club Sales\",SalesDepartment.Description)"
    GroupLevel =0
    Expression ="InvoLine5.TransType"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TotalQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DepartmentDesc"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Department"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.TransType"
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
    Bottom =229
    Left =0
    Top =16
    ColumnsShown =543
    Begin
        Left =608
        Top =-15
        Right =704
        Bottom =99
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =476
        Top =16
        Right =572
        Bottom =130
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =201
        Top =-5
        Right =297
        Bottom =214
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =341
        Top =20
        Right =437
        Bottom =134
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =39
        Top =18
        Right =135
        Bottom =132
        Top =0
        Name ="Cashier"
        Name =""
    End
End
