Operation =1
Option =0
Where ="(((ItemLabels.Labels)<>0))"
Begin InputTables
    Name ="ItemLabels"
    Name ="qryUPCGlyph"
    Name ="Size"
End
Begin OutputColumns
    Expression ="qryUPCGlyph.UPCG"
    Alias ="BarCode"
    Expression ="IIf(IsNull([ItemLabels].[UPCCodeMax]) Or [ItemLabels].[ID]<1000,[ItemLabels].[ID"
        "],[ItemLabels].[UPCCodeMax])"
    Alias ="Descr"
    Expression ="ItemLabels.ItemD"
    Expression ="ItemLabels.ID"
    Expression ="ItemLabels.Labels"
    Expression ="ItemLabels.SizeD"
    Expression ="ItemLabels.CurPrice"
    Expression ="ItemLabels.S"
    Expression ="ItemLabels.SalesDpt"
    Expression ="ItemLabels.Price"
    Expression ="ItemLabels.VendorName"
    Expression ="ItemLabels.SetupDate"
    Expression ="ItemLabels.ChangeDate"
    Alias ="POZ"
    Expression ="IIf([oz]>0,Format([CurPrice]/[Oz],\"#.00\"\" /oz\"\"\"),\"\")"
End
Begin Joins
    LeftTable ="ItemLabels"
    RightTable ="qryUPCGlyph"
    Expression ="ItemLabels.ID=qryUPCGlyph.ItemID"
    Flag =2
    LeftTable ="ItemLabels"
    RightTable ="Size"
    Expression ="ItemLabels.SizeD=Size.Description"
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
        dbText "Name" ="Descr"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryUPCGlyph.UPCG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.Labels"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.SizeD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.CurPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.S"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.SalesDpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemLabels.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BarCode"
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
    Bottom =149
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =28
        Top =12
        Right =174
        Bottom =201
        Top =0
        Name ="ItemLabels"
        Name =""
    End
    Begin
        Left =216
        Top =12
        Right =401
        Bottom =109
        Top =0
        Name ="qryUPCGlyph"
        Name =""
    End
    Begin
        Left =240
        Top =113
        Right =414
        Bottom =227
        Top =0
        Name ="Size"
        Name =""
    End
End
