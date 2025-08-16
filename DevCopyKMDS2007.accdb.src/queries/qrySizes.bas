Operation =1
Option =0
Begin InputTables
    Name ="qryItemsBySizeCounts"
    Name ="Size"
End
Begin OutputColumns
    Expression ="Size.ID"
    Expression ="Size.Description"
    Expression ="Size.Oz"
    Expression ="Size.CaseQty"
    Alias ="Items"
    Expression ="IIf(IsNull([qryItemsBySizeCounts].[count]),0,[qryItemsBySizeCounts].[count])"
End
Begin Joins
    LeftTable ="qryItemsBySizeCounts"
    RightTable ="Size"
    Expression ="qryItemsBySizeCounts.SizeID = Size.ID"
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
        dbText "Name" ="Size.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Items"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size.Oz"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Size.CaseQty"
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
    Bottom =150
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =182
        Top =12
        Right =373
        Bottom =156
        Top =0
        Name ="qryItemsBySizeCounts"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Size"
        Name =""
    End
End
