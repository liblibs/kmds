Operation =1
Option =0
Having ="(((InvoLine.Code)=[parCode]))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="Count"
    Expression ="Sum(1)"
End
Begin Parameters
    Name ="parCode"
    Flag =16
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
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
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-340
    Top =55
    Right =885
    Bottom =732
    Left =-1
    Top =-1
    Right =1193
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
