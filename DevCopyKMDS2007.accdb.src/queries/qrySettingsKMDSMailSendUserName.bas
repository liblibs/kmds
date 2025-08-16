Operation =1
Option =0
Where ="(((Settings.KeyName)=\"KMDSMail sendusername\"))"
Begin InputTables
    Name ="Settings"
End
Begin OutputColumns
    Expression ="Settings.KeyName"
    Expression ="Settings.Value"
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
        dbText "Name" ="Settings.Value"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =412
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =97
        Top =21
        Right =241
        Bottom =165
        Top =0
        Name ="Settings"
        Name =""
    End
End
