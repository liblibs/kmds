Operation =1
Option =0
Where ="(((ItemLabels.Labels)>0))"
Begin InputTables
    Name ="ItemLabels"
    Name ="qryUPCGlyph"
End
Begin OutputColumns
    Alias ="UPCCode"
    Expression ="qryUPCGlyph.ID"
    Alias ="Descr"
    Expression ="Left([ItemD],20)"
    Expression ="qryUPCGlyph.ItemID"
    Expression ="ItemLabels.SalesDpt"
    Expression ="ItemLabels.CurPrice"
    Expression ="qryUPCGlyph.UPCG"
    Expression ="ItemLabels.S"
    Expression ="ItemLabels.Labels"
End
Begin Joins
    LeftTable ="ItemLabels"
    RightTable ="qryUPCGlyph"
    Expression ="ItemLabels.ID = qryUPCGlyph.ItemID"
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
        dbText "Name" ="UPCCode"
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
    Bottom =223
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =25
        Top =85
        Right =152
        Bottom =235
        Top =0
        Name ="ItemLabels"
        Name =""
    End
    Begin
        Left =375
        Top =-10
        Right =560
        Bottom =132
        Top =0
        Name ="qryUPCGlyph"
        Name =""
    End
End
