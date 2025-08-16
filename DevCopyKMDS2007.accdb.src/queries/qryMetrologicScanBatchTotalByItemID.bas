Operation =1
Option =0
Begin InputTables
    Name ="MetrologicScanbatch"
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ItemID"
    Alias ="SumOfQuantity"
    Expression ="Sum(MetrologicScanbatch.Quantity)"
    Alias ="FirstOfCashierID"
    Expression ="First(MetrologicScanbatch.CashierID)"
    Alias ="InvoLineID"
    Expression ="First(MetrologicScanbatch.ID)"
End
Begin Joins
    LeftTable ="MetrologicScanbatch"
    RightTable ="UPCItem"
    Expression ="MetrologicScanbatch.ScanCode=UPCItem.ID"
    Flag =1
End
Begin Groups
    Expression ="UPCItem.ItemID"
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
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfCashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLineID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1509
    Bottom =851
    Left =-1
    Top =-1
    Right =1401
    Bottom =402
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =283
        Top =0
        Name ="MetrologicScanbatch"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
