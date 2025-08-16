Operation =1
Option =0
Where ="(((Settings.KeyName)=\"Fintech.Division-IDRelated\"))"
Begin InputTables
    Name ="Settings"
End
Begin OutputColumns
    Alias ="Division_ID"
    Expression ="CLng([Value])"
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
        dbText "Name" ="Division_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =120
    Right =1531
    Bottom =857
    Left =-1
    Top =-1
    Right =1457
    Bottom =385
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Settings"
        Name =""
    End
End
