Operation =5
Option =0
Begin InputTables
    Name ="RecreateInventoryTransmit"
End
Begin OutputColumns
    Expression ="RecreateInventoryTransmit.*"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Movement"
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
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="RecreateInventoryTransmit"
        Name =""
    End
End
