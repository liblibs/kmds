Operation =1
Option =0
Where ="(((OLCC_8_and_12_UPC.Code)<>[UPCItem].[itemid]))"
Begin InputTables
    Name ="OLCC_8_and_12_UPC"
    Name ="UPCItem"
End
Begin OutputColumns
    Expression ="OLCC_8_and_12_UPC.Code"
    Expression ="UPCItem.ItemID"
    Expression ="OLCC_8_and_12_UPC.upc12"
    Expression ="OLCC_8_and_12_UPC.UPCCode"
End
Begin Joins
    LeftTable ="OLCC_8_and_12_UPC"
    RightTable ="UPCItem"
    Expression ="OLCC_8_and_12_UPC.upc12 = UPCItem.ID"
    Flag =1
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
        dbText "Name" ="OLCC_8_and_12_UPC.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCC_8_and_12_UPC.upc12"
        dbInteger "ColumnWidth" ="2325"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCC_8_and_12_UPC.UPCCode"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =54
    Top =96
    Right =1216
    Bottom =766
    Left =-1
    Top =-1
    Right =1139
    Bottom =400
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =32
        Top =12
        Right =176
        Bottom =239
        Top =0
        Name ="OLCC_8_and_12_UPC"
        Name =""
    End
    Begin
        Left =288
        Top =59
        Right =432
        Bottom =203
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
