Operation =1
Option =0
Where ="(((dbo_OLCCScanCode.Code)<>[upcitem].[ItemID]))"
Begin InputTables
    Name ="dbo_OLCCScanCode"
    Name ="Item"
    Name ="Item"
    Alias ="Item_1"
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="UPCItem.ID"
    Alias ="KMDSItem"
    Expression ="UPCItem.ItemID"
    Alias ="KMDSDesc"
    Expression ="Item_1.Description"
    Expression ="Item_1.Onhand"
    Alias ="OLCCItem"
    Expression ="dbo_OLCCScanCode.Code"
    Alias ="OLCCDesc"
    Expression ="Item.Description"
    Expression ="dbo_OLCCScanCode.SetupDate"
End
Begin Joins
    LeftTable ="dbo_OLCCScanCode"
    RightTable ="UPCItem"
    Expression ="dbo_OLCCScanCode.UPCCodeKMDS = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="Item_1"
    Expression ="UPCItem.ItemID = Item_1.ID"
    Flag =1
    LeftTable ="dbo_OLCCScanCode"
    RightTable ="Item"
    Expression ="dbo_OLCCScanCode.Code = Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="Item_1.Description"
    Flag =0
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
        dbText "Name" ="Item_1.Onhand"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSItem"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSDesc"
        dbInteger "ColumnWidth" ="3135"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCDesc"
        dbInteger "ColumnWidth" ="3180"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_OLCCScanCode.SetupDate"
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
    Bottom =294
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="dbo_OLCCScanCode"
        Name =""
    End
    Begin
        Left =442
        Top =10
        Right =586
        Bottom =154
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =441
        Top =163
        Right =585
        Bottom =319
        Top =0
        Name ="Item_1"
        Name =""
    End
    Begin
        Left =221
        Top =171
        Right =365
        Bottom =315
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
