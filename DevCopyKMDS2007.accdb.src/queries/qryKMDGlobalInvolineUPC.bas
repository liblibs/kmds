Operation =1
Option =0
Where ="(((Len([InputCode]))>5))"
Having ="(((InvoLine.Code) Between 1000 And 99999))"
Begin InputTables
    Name ="InvoLine"
    Name ="UPCItem"
    Name ="Sysfile"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Expression ="InvoLine.InputCode"
    Alias ="LastOfSetupDate"
    Expression ="Last(InvoLine.SetupDate)"
    Expression ="Sysfile.AgentNumber"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="UPCItem"
    Expression ="InvoLine.InputCode = UPCItem.ID"
    Flag =1
End
Begin Groups
    Expression ="InvoLine.Code"
    GroupLevel =0
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="Sysfile.AgentNumber"
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
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="3345"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastOfSetupDate"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sysfile.AgentNumber"
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
    Bottom =332
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =391
        Bottom =358
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =439
        Top =12
        Right =583
        Bottom =156
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =439
        Top =156
        Right =583
        Bottom =300
        Top =0
        Name ="Sysfile"
        Name =""
    End
End
