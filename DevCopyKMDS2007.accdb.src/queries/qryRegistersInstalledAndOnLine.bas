Operation =1
Option =0
Where ="(((CashSys.ID) Between 1 And 5))"
Begin InputTables
    Name ="Sysfile"
    Name ="CashSys"
End
Begin OutputColumns
    Expression ="CashSys.ID"
    Expression ="CashSys.ComputerName"
    Expression ="Sysfile.Cr1Installed"
    Expression ="Sysfile.Cr2Installed"
    Expression ="Sysfile.Cr3Installed"
    Expression ="Sysfile.Cr4Installed"
    Expression ="Sysfile.Cr1Online"
    Expression ="Sysfile.Cr2Online"
    Expression ="Sysfile.Cr3Online"
    Expression ="Sysfile.Cr4Online"
    Expression ="Sysfile.Cr5Installed"
    Expression ="Sysfile.Cr5Online"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CashSys.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr5Installed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.Cr5Online"
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
    Bottom =317
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =303
        Bottom =321
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="CashSys"
        Name =""
    End
End
