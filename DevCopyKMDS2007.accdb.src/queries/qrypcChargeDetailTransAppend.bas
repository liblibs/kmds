Operation =3
Name ="pcChargeDebitsAndCredits"
Option =0
Where ="(((pcChargeTrans.Station)=[parUser]) And ((pcChargeTrans.Date)=[parDate]))"
Begin InputTables
    Name ="pcChargeTrans"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="Station"
    Expression ="pcChargeTrans.Station"
    Alias ="Expr2"
    Name ="Date"
    Expression ="pcChargeTrans.Date"
    Alias ="Expr3"
    Name ="Time"
    Expression ="pcChargeTrans.Time"
    Alias ="Expr4"
    Name ="Number"
    Expression ="pcChargeTrans.Number"
    Alias ="Expr5"
    Name ="Card"
    Expression ="pcChargeTrans.Card"
    Alias ="Expr6"
    Name ="Member"
    Expression ="pcChargeTrans.Member"
    Alias ="Expr7"
    Name ="Result"
    Expression ="pcChargeTrans.Result"
    Alias ="Expr8"
    Name ="Auth"
    Expression ="pcChargeTrans.Auth"
    Alias ="Expr9"
    Name ="Amount"
    Expression ="pcChargeTrans.Amount"
    Alias ="CapturedAmount"
    Name ="CapturedAmount"
    Expression ="IIf([Result]=\"CAPTURED\",[Total_Auth],0)"
    Alias ="DC"
    Name ="DebitOrCredit"
    Expression ="\"C\""
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
    Expression ="pcChargeTrans.Date"
    Flag =0
    Expression ="pcChargeTrans.Time"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
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
        dbText "Name" ="CapturedAmount"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="pcChargeTrans"
        Name =""
    End
End
