Operation =1
Option =0
Where ="(((InvoLine.TransType)=\"Y\"))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Alias ="MaxOfSetupDate"
    Expression ="Max(InvoLine.SetupDate)"
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
        dbText "Name" ="MaxOfSetupDate"
        dbInteger "ColumnWidth" ="4230"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =298
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =343
        Bottom =357
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
