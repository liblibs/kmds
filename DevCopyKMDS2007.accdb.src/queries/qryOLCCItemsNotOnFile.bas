Operation =1
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="dbo_OLCCitem"
    Name ="Item"
End
Begin OutputColumns
    Expression ="dbo_OLCCitem.Code"
    Expression ="dbo_OLCCitem.Description"
    Expression ="dbo_OLCCitem.Price"
    Expression ="dbo_OLCCitem.UnitsPerCase"
    Expression ="dbo_OLCCitem.StockStatus"
    Alias ="SizeID"
    Expression ="Right([Code],1)"
End
Begin Joins
    LeftTable ="dbo_OLCCitem"
    RightTable ="Item"
    Expression ="dbo_OLCCitem.Code = Item.ID"
    Flag =2
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
        dbText "Name" ="dbo_OLCCitem.Code"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Description"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCitem.StockStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
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
    Right =1193
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =253
        Bottom =273
        Top =0
        Name ="dbo_OLCCitem"
        Name =""
    End
    Begin
        Left =381
        Top =8
        Right =525
        Bottom =152
        Top =0
        Name ="Item"
        Name =""
    End
End
