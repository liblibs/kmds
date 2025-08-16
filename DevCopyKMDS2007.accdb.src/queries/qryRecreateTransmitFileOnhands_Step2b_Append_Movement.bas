Operation =3
Name ="RecreateInventoryTransmit"
Option =0
Where ="(((InvoLine.TransType)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransTyp"
    "e)=\"P\" Or (InvoLine.TransType)=\"X\" Or (InvoLine.TransType)=\"A\") And ((Invo"
    "Line.SetupDate)>Forms!ReportsExplorer!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Name ="Code"
    Expression ="InvoLine.Code"
    Alias ="Movement"
    Name ="Movement"
    Expression ="Sum(IIf([TransType]=\"S\" Or [TransType]=\"C\",[Quantity],0-[Quantity]))"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Movement"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-25
    Top =31
    Right =1200
    Bottom =708
    Left =-1
    Top =-1
    Right =1193
    Bottom =232
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =383
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
