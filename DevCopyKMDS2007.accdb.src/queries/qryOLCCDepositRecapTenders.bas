Operation =1
Option =0
Where ="(((InvoLine5.TransType)=\"T\") AND ((InvoLine5.InputCode) Like \"Drop*\" Or (Inv"
    "oLine5.InputCode)=\"Cash\" Or (InvoLine5.InputCode)=\"Fintech\"))"
Having ="(((0-Sum(IIf([Code]=2,[Price],0)))<>1))"
Begin InputTables
    Name ="InvoLine5"
End
Begin OutputColumns
    Alias ="Checks"
    Expression ="0-Sum(IIf([Code]=2,[Price],0))"
    Alias ="Cards"
    Expression ="0-Sum(IIf([Code]=3,[Price],0))"
    Alias ="EFT"
    Expression ="0-Sum(IIf([Code]=10,[Price],0))"
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
        dbText "Name" ="Checks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cards"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EFT"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-15
    Top =201
    Right =1210
    Bottom =878
    Left =-1
    Top =-1
    Right =1193
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine5"
        Name =""
    End
End
