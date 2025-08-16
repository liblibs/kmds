Operation =2
Name ="InvoLineUPCMaxDate"
Option =0
Having ="(((InvoLine.TransType)=\"S\") AND ((Len([INPUTCODE]))>7))"
Begin InputTables
    Name ="InvoLine"
End
Begin OutputColumns
    Expression ="InvoLine.InputCode"
    Alias ="MaxOfSetupDate"
    Expression ="Max(InvoLine.SetupDate)"
End
Begin Groups
    Expression ="InvoLine.InputCode"
    GroupLevel =0
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="Len([INPUTCODE])"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="MaxOfSetupDate"
        dbInteger "ColumnWidth" ="3105"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =215
        Top =6
        Right =311
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
End
