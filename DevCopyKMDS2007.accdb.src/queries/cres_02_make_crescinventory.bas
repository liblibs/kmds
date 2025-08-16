Operation =2
Name ="CrescInventory"
Option =0
Where ="(((InvoLine.SetupDate)>#3/9/2017#) AND ((InvoLine.TransType)=\"S\"))"
Having ="(((Item.SalesDepartmentID)>21))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.SalesDepartmentID"
    Alias ="SumOfQuantity"
    Expression ="Sum(InvoLine.Quantity)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item.ID"
    Flag =0
End
Begin Groups
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.SalesDepartmentID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =71
    Right =1575
    Bottom =904
    Left =-1
    Top =-1
    Right =1508
    Bottom =531
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =421
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =323
        Top =0
        Name ="Item"
        Name =""
    End
End
