Operation =1
Option =0
Begin InputTables
    Name ="KMDSKS"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="Sysfile.AgentNumber"
    Expression ="Sysfile.Name"
    Expression ="KMDSKS.SetupDate"
End
Begin Joins
    LeftTable ="Sysfile"
    RightTable ="KMDSKS"
    Expression ="Sysfile.AgentNumber = KMDSKS.AgentNumber"
    Flag =1
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
        dbText "Name" ="Sysfile.AgentNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSKS.SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =88
    Right =1025
    Bottom =846
    Left =-1
    Top =-1
    Right =954
    Bottom =474
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="KMDSKS"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
