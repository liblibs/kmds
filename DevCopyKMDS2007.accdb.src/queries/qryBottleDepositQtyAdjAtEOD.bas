Operation =1
Option =0
Where ="(((Settings.KeyName)=\"BottleDepositQtyAdjAtEOD\") AND ((UCase([Value]))=\"TRUE\""
    "))"
Begin InputTables
    Name ="Settings"
End
Begin OutputColumns
    Expression ="Settings.KeyName"
    Alias ="UV"
    Expression ="UCase([Value])"
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
        dbText "Name" ="Settings.KeyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UV"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-336
    Top =132
    Right =2824
    Bottom =1940
    Left =-1
    Top =-1
    Right =3122
    Bottom =1433
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="Settings"
        Name =""
    End
End
