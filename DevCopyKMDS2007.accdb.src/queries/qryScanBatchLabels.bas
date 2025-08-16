Operation =1
Option =0
Begin InputTables
    Name ="ScanBatchQuantityLabels"
End
Begin OutputColumns
    Expression ="ScanBatchQuantityLabels.UPCCode"
    Alias ="UPCG"
    Expression ="UPCGlyph([UPCCode])"
    Expression ="ScanBatchQuantityLabels.Description"
    Alias ="F39"
    Expression ="\"*\" & IIf(CLng([UPCCode])<10,\"0\") & CLng([UPCCode]) & \"*\""
    Alias ="F39x"
    Expression ="\"* \" & CLng([UPCCode]) & \"*\""
End
Begin OrderBy
    Expression ="ScanBatchQuantityLabels.UPCCode"
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
        dbText "Name" ="ScanBatchQuantityLabels.UPCCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCG"
        dbInteger "ColumnWidth" ="3570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScanBatchQuantityLabels.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="F39"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="F39x"
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
    Bottom =135
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =290
        Bottom =214
        Top =0
        Name ="ScanBatchQuantityLabels"
        Name =""
    End
End
