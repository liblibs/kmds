Operation =1
Option =0
Begin InputTables
    Name ="pcChargeDebitsAndCredits"
End
Begin OutputColumns
    Expression ="pcChargeDebitsAndCredits.Station"
    Expression ="pcChargeDebitsAndCredits.Date"
    Expression ="pcChargeDebitsAndCredits.Time"
    Expression ="pcChargeDebitsAndCredits.Number"
    Expression ="pcChargeDebitsAndCredits.Card"
    Expression ="pcChargeDebitsAndCredits.Member"
    Expression ="pcChargeDebitsAndCredits.Result"
    Expression ="pcChargeDebitsAndCredits.Auth"
    Expression ="pcChargeDebitsAndCredits.Amount"
    Expression ="pcChargeDebitsAndCredits.CapturedAmount"
    Alias ="DC"
    Expression ="IIf([Result]=\"Captured\",[DebitOrCredit],\"-\")"
End
Begin OrderBy
    Expression ="pcChargeDebitsAndCredits.Date"
    Flag =0
    Expression ="pcChargeDebitsAndCredits.Time"
    Flag =0
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
        dbText "Name" ="pcChargeDebitsAndCredits.Result"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DC"
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
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =258
        Bottom =120
        Top =0
        Name ="pcChargeDebitsAndCredits"
        Name =""
    End
End
