Operation =1
Option =0
Having ="(((Len([ID]))=12) AND ((Sum(1))=2))"
Begin InputTables
    Name ="UPCItem"
End
Begin OutputColumns
    Alias ="LID"
    Expression ="Len([ID])"
    Alias ="COUNT"
    Expression ="Sum(1)"
    Alias ="ID10"
    Expression ="Mid([ID],2,10)"
    Alias ="cd"
    Expression ="checkdigit11(\"0\" & Mid([ID],2,10))"
End
Begin Groups
    Expression ="Len([ID])"
    GroupLevel =0
    Expression ="Mid([ID],2,10)"
    GroupLevel =0
    Expression ="checkdigit11(\"0\" & Mid([ID],2,10))"
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
        dbText "Name" ="LID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COUNT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cd"
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
