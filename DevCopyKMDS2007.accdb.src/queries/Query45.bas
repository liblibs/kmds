Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="olcc order"
    Name ="OLCCitem"
End
Begin OutputColumns
    Expression ="[olcc order].[ITEM ]"
    Expression ="[olcc order].DESCRIPTION"
    Expression ="[olcc order].[ORDER ]"
    Expression ="[olcc order].ID"
    Expression ="Item.ID"
    Expression ="OLCCitem.Code"
    Expression ="OLCCitem.Description"
End
Begin Joins
    LeftTable ="olcc order"
    RightTable ="Item"
    Expression ="[olcc order].[ITEM ] = Item.ID"
    Flag =2
    LeftTable ="olcc order"
    RightTable ="OLCCitem"
    Expression ="[olcc order].[ITEM ] = OLCCitem.Code"
    Flag =2
End
Begin OrderBy
    Expression ="Item.ID"
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
        dbText "Name" ="[olcc order].[ITEM ]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[olcc order].DESCRIPTION"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[olcc order].[ORDER ]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[olcc order].ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OLCCitem.Description"
        dbInteger "ColumnWidth" ="4815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1531
    Bottom =873
    Left =-1
    Top =-1
    Right =1508
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =649
        Top =22
        Right =793
        Bottom =166
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="olcc order"
        Name =""
    End
    Begin
        Left =913
        Top =288
        Right =1057
        Bottom =536
        Top =0
        Name ="OLCCitem"
        Name =""
    End
End
