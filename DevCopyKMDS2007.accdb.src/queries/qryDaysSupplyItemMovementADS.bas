Operation =1
Option =0
Where ="(((ItemMovement.Description)=\"Sales - Total\"))"
Begin InputTables
    Name ="ItemMovement"
End
Begin OutputColumns
    Expression ="ItemMovement.Description"
    Expression ="ItemMovement.Code"
    Alias ="ADSCurrMM"
    Expression ="[CurrMM]/Day(Now())"
    Alias ="ADSJan"
    Expression ="[Jan]/31"
    Alias ="ADSFeb"
    Expression ="[Feb]/28"
    Alias ="ADSMar"
    Expression ="[Mar]/31"
    Alias ="ADSApr"
    Expression ="[Apr]/30"
    Alias ="ADSMay"
    Expression ="[May]/31"
    Alias ="ADSJun"
    Expression ="[Jun]/30"
    Alias ="ADSJul"
    Expression ="[Jul]/31"
    Alias ="ADSAug"
    Expression ="[Aug]/31"
    Alias ="ADSSep"
    Expression ="[Sep]/30"
    Alias ="ADSOct"
    Expression ="[Oct]/31"
    Alias ="ADSNov"
    Expression ="[Nov]/30"
    Alias ="ADSDec"
    Expression ="[Dec]/31"
    Alias ="today"
    Expression ="Day(Now())"
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
        dbText "Name" ="ADSCurrMM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSFeb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSMar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSApr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSMay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSJun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSJul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSAug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSSep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSOct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSNov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSDec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="today"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ADSJan"
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
        Name ="ItemMovement"
        Name =""
    End
End
