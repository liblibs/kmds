Operation =1
Option =0
Where ="(((Item.SubCategoryID) Between [Forms]![LabelList]![cboCategoryFrom] And [Forms]"
    "![LabelList]![cboCategoryTo]))"
Begin InputTables
    Name ="ItemLabels"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SubCategoryID"
    Expression ="ItemLabels.ItemD"
    Alias ="Price"
    Expression ="ItemLabels.curprice"
    Expression ="ItemLabels.ID"
End
Begin Joins
    LeftTable ="ItemLabels"
    RightTable ="Item"
    Expression ="ItemLabels.ID=Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="ItemLabels.ItemD"
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
        dbText "Name" ="Item.SubCategoryID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.ItemD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
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
    Bottom =226
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =266
        Top =0
        Name ="ItemLabels"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =234
        Top =0
        Name ="Item"
        Name =""
    End
End
