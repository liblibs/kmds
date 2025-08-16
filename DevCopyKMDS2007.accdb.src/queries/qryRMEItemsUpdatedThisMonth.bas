Operation =1
Option =0
Where ="(((InvoLine.SetupDate)>Month(Now()) & \"/01/\" & Year(Now())) AND ((Item.SalesDe"
    "partmentID)=1))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="Counter"
    Expression ="Sum(IIf([Transtype]=\"S\",[Quantity],0))"
    Alias ="Clubs"
    Expression ="Sum(IIf([Transtype]=\"C\",[Quantity],0))"
    Alias ="Purchase"
    Expression ="Sum(IIf([Transtype]=\"P\",[Quantity],0))"
    Alias ="OnhandChange"
    Expression ="Sum(IIf([Transtype]=\"A\",[Quantity],0))"
    Alias ="Transfer"
    Expression ="Sum(IIf([Transtype]=\"X\",[Quantity],0))"
    Alias ="Delta"
    Expression ="[Counter]+[Clubs]-[Purchase]-[Transfer]-[OnhandChange]"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="InvoLine.Code"
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
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Counter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Purchase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnhandChange"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transfer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delta"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =14
    Top =197
    Right =883
    Bottom =890
    Left =-1
    Top =-1
    Right =837
    Bottom =375
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =272
        Bottom =373
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =329
        Top =43
        Right =473
        Bottom =187
        Top =0
        Name ="Item"
        Name =""
    End
End
