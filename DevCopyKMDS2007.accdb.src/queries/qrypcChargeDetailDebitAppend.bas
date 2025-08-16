Operation =3
Name ="pcChargeDebitsAndCredits"
Option =0
Where ="(((pcChargeDebit.Station)=[parUser]) And ((pcChargeDebit.Date)=[parDate]))"
Begin InputTables
    Name ="pcChargeDebit"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="Station"
    Expression ="pcChargeDebit.Station"
    Alias ="Expr2"
    Name ="Date"
    Expression ="pcChargeDebit.Date"
    Alias ="Expr3"
    Name ="Time"
    Expression ="pcChargeDebit.Time"
    Alias ="Expr4"
    Name ="Number"
    Expression ="pcChargeDebit.Number"
    Alias ="Expr5"
    Name ="Card"
    Expression ="pcChargeDebit.Card"
    Alias ="Expr6"
    Name ="Member"
    Expression ="pcChargeDebit.Member"
    Alias ="Expr7"
    Name ="Result"
    Expression ="pcChargeDebit.Result"
    Alias ="Expr8"
    Name ="Auth"
    Expression ="pcChargeDebit.Auth"
    Alias ="Expr9"
    Name ="Amount"
    Expression ="pcChargeDebit.Amount"
    Alias ="CapturedAmount"
    Name ="CapturedAmount"
    Expression ="IIf([Result]=\"CAPTURED\",[Amount],0)"
    Alias ="DC"
    Name ="DebitOrCredit"
    Expression ="\"D\""
    Alias ="TicketNumber"
    Name ="TicketNumber"
    Expression ="IIf(IsNull([Ticket]) Or [Ticket]=\"\",0,[Ticket])"
End
Begin Parameters
    Name ="parUser"
    Flag =10
    Name ="parDate"
    Flag =8
End
Begin OrderBy
    Expression ="pcChargeDebit.Date"
    Flag =0
    Expression ="pcChargeDebit.Time"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="pcChargeDebit.Station"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Number"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Date"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Time"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Card"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Member"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Result"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Auth"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pcChargeDebit.Amount"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CapturedAmount"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TicketNumber"
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
    Bottom =334
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =268
        Bottom =330
        Top =0
        Name ="pcChargeDebit"
        Name =""
    End
End
