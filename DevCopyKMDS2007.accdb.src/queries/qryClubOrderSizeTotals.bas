Operation =1
Option =0
Begin InputTables
    Name ="qryClubOrderSizeTot"
End
Begin OutputColumns
    Expression ="qryClubOrderSizeTot.Description"
    Expression ="qryClubOrderSizeTot.Count"
    Expression ="qryClubOrderSizeTot.CQ"
    Expression ="qryClubOrderSizeTot.QS"
    Alias ="Full"
    Expression ="[QS]\\[CQ]"
    Alias ="Broken"
    Expression ="[QS]-(([QS]\\[CQ])*[CQ])"
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
        dbText "Name" ="Full"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Broken"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="qryClubOrderSizeTot"
        Name =""
    End
End
