Operation =1
Option =0
Where ="(((InvoLine.SetupDate) Between [Forms]![PrintInventoryReport]![txtDateRangeFrom]"
    " And [Forms]![PrintInventoryReport]![txtDateRangeTo]))"
Having ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransTyp"
    "e)=\"D\"))"
Begin InputTables
    Name ="InvoLine"
    Name ="SalesGroup"
    Name ="SalesDepartment"
    Name ="Item"
    Name ="Top150"
End
Begin OutputColumns
    Expression ="SalesGroup.GroupLevel"
    Alias ="Group"
    Expression ="SalesGroup.ID"
    Alias ="GroupDesc"
    Expression ="SalesGroup.Description"
    Alias ="Department"
    Expression ="SalesDepartment.ID"
    Alias ="DepartmentDesc"
    Expression ="IIf(involine.Transtype='C',\"Club Sales\",SalesDepartment.Description)"
    Alias ="TotalQ"
    Expression ="Sum(InvoLine.Quantity)"
    Alias ="TotalP"
    Expression ="Sum(InvoLine.Price)"
    Alias ="Q1"
    Expression ="Sum(IIf(IsNull([Top150]![Code]),[Quantity],0))"
    Alias ="D1"
    Expression ="Sum(IIf(IsNull([Top150]![Code]),[involine].[Price],0))"
    Alias ="Q2"
    Expression ="Sum(IIf(IsNull([Top150]![Code]),0,[Quantity]))"
    Alias ="D2"
    Expression ="Sum(IIf(IsNull([Top150]![Code]),0,[involine].[Price]))"
    Alias ="Q3"
    Expression ="Sum(IIf([CashsysID]=3,[Quantity],0))"
    Alias ="D3"
    Expression ="Sum(IIf([CashsysID]=3,involine.Price,0))"
    Alias ="Q4"
    Expression ="Sum(IIf([CashsysID]=4,[Quantity],0))"
    Alias ="D4"
    Expression ="Sum(IIf([CashsysID]=4,involine.Price,0))"
    Expression ="InvoLine.TransType"
    Alias ="Q5"
    Expression ="Sum(IIf([CashsysID]=5,[Quantity],0))"
    Alias ="D5"
    Expression ="Sum(IIf([CashsysID]=5,[involine].[Price],0))"
End
Begin Joins
    LeftTable ="SalesDepartment"
    RightTable ="Item"
    Expression ="SalesDepartment.ID = Item.SalesDepartmentID"
    Flag =1
    LeftTable ="SalesGroup"
    RightTable ="SalesDepartment"
    Expression ="SalesGroup.ID = SalesDepartment.SalesGroupID"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="Top150"
    Expression ="Item.ID = Top150.Code"
    Flag =2
End
Begin OrderBy
    Expression ="SalesGroup.ID"
    Flag =0
    Expression ="SalesDepartment.ID"
    Flag =0
End
Begin Groups
    Expression ="SalesGroup.GroupLevel"
    GroupLevel =0
    Expression ="SalesGroup.ID"
    GroupLevel =0
    Expression ="SalesGroup.Description"
    GroupLevel =0
    Expression ="SalesDepartment.ID"
    GroupLevel =0
    Expression ="IIf(involine.Transtype='C',\"Club Sales\",SalesDepartment.Description)"
    GroupLevel =0
    Expression ="InvoLine.TransType"
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
        dbText "Name" ="SalesGroup.GroupLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1402
    Bottom =677
    Left =-1
    Top =-1
    Right =1370
    Bottom =317
    Left =0
    Top =16
    ColumnsShown =543
    Begin
        Left =129
        Top =29
        Right =253
        Bottom =218
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =660
        Top =37
        Right =756
        Bottom =151
        Top =0
        Name ="SalesGroup"
        Name =""
    End
    Begin
        Left =488
        Top =14
        Right =584
        Bottom =128
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =333
        Top =31
        Right =432
        Bottom =297
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =514
        Top =138
        Right =658
        Bottom =282
        Top =0
        Name ="Top150"
        Name =""
    End
End
