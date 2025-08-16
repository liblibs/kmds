Operation =6
Option =0
Begin InputTables
    Name ="CashiersAndPrograms"
End
Begin OutputColumns
    Expression ="CashiersAndPrograms.ProgramName"
    GroupLevel =2
    Expression ="CashiersAndPrograms.CashierID"
    GroupLevel =1
    Alias ="FirstOfAllowed"
    Expression ="First(CashiersAndPrograms.Allowed)"
    Alias ="Total Of Allowed"
    Expression ="First(CashiersAndPrograms.Allowed)"
    GroupLevel =2
End
Begin Groups
    Expression ="CashiersAndPrograms.ProgramName"
    GroupLevel =2
    Expression ="CashiersAndPrograms.CashierID"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="Total Of Allowed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashiersAndPrograms.ProgramName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="99"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="27"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="CashiersAndPrograms"
        Name =""
    End
End
