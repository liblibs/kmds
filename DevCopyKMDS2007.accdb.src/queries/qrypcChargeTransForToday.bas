Operation =1
Option =0
Where ="((([pcChargeTrans].[Date]) Between Format(Now(),\"mm/dd/yy\"\" 00\"\":\"\"00\"\""
    "\") And Format(Now(),\"mm/dd/yyyy\"\" 23\"\":\"\"59\"\":\"\"59\"\"\")))"
Begin InputTables
    Name ="pcChargeTrans"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="pcChargeTrans.Date"
    Alias ="Expr2"
    Expression ="pcChargeTrans.Number"
    Alias ="Expr3"
    Expression ="pcChargeTrans.Ticket"
    Alias ="Expr4"
    Expression ="pcChargeTrans.Date"
    Alias ="Expr5"
    Expression ="pcChargeTrans.Time"
    Alias ="Expr6"
    Expression ="pcChargeTrans.Station"
    Alias ="Expr7"
    Expression ="pcChargeTrans.Processor"
    Alias ="Expr8"
    Expression ="pcChargeTrans.TID"
    Alias ="Expr9"
    Expression ="pcChargeTrans.Card"
    Alias ="Expr10"
    Expression ="pcChargeTrans.Issuer"
    Alias ="Expr11"
    Expression ="pcChargeTrans.Member"
    Alias ="Expr12"
    Expression ="pcChargeTrans.ExpDate"
    Alias ="Expr13"
    Expression ="pcChargeTrans.Action"
    Alias ="Expr14"
    Expression ="pcChargeTrans.Manual"
    Alias ="Expr15"
    Expression ="pcChargeTrans.Amount"
    Alias ="Expr16"
    Expression ="pcChargeTrans.Ref"
    Alias ="Expr17"
    Expression ="pcChargeTrans.Zip"
    Alias ="Expr18"
    Expression ="pcChargeTrans.Street"
    Alias ="Expr19"
    Expression ="pcChargeTrans.Result"
    Alias ="Expr20"
    Expression ="pcChargeTrans.Auth"
    Alias ="Expr21"
    Expression ="pcChargeTrans.Result_Ref"
    Alias ="Expr22"
    Expression ="pcChargeTrans.Result_Add"
    Alias ="Expr23"
    Expression ="pcChargeTrans.[CID Number]"
    Alias ="Expr24"
    Expression ="pcChargeTrans.Response"
    Alias ="Expr25"
    Expression ="pcChargeTrans.[Tax Amount]"
    Alias ="Expr26"
    Expression ="pcChargeTrans.Post_Amount"
    Alias ="Expr27"
    Expression ="pcChargeTrans.Total_Auth"
    Alias ="Expr28"
    Expression ="pcChargeTrans.IND"
    Alias ="Expr29"
    Expression ="pcChargeTrans.TranID"
    Alias ="Expr30"
    Expression ="pcChargeTrans.ValCode"
    Alias ="Expr31"
    Expression ="pcChargeTrans.MSI"
    Alias ="Expr32"
    Expression ="pcChargeTrans.RET"
    Alias ="Expr33"
    Expression ="pcChargeTrans.PEM"
    Alias ="Expr34"
    Expression ="pcChargeTrans.Trans_Indicator"
    Alias ="Expr35"
    Expression ="pcChargeTrans.ReqACI"
    Alias ="Expr36"
    Expression ="pcChargeTrans.RetACI"
    Alias ="Expr37"
    Expression ="pcChargeTrans.RespCode"
    Alias ="Expr38"
    Expression ="pcChargeTrans.DownGrade"
    Alias ="Expr39"
    Expression ="pcChargeTrans.TransDate"
    Alias ="Expr40"
    Expression ="pcChargeTrans.TransTime"
    Alias ="Expr41"
    Expression ="pcChargeTrans.EncryptedCard"
    Alias ="Expr42"
    Expression ="pcChargeTrans.PeriodicPayment"
    Alias ="Expr43"
    Expression ="pcChargeTrans.BatchNumber"
    Alias ="Expr44"
    Expression ="pcChargeTrans.ItemNumber"
    Alias ="Expr45"
    Expression ="pcChargeTrans.[Business Date]"
    Alias ="Expr46"
    Expression ="pcChargeTrans.Gratuity"
    Alias ="Expr47"
    Expression ="pcChargeTrans.BType"
    Alias ="Expr48"
    Expression ="pcChargeTrans.ServerID"
    Alias ="Expr49"
    Expression ="pcChargeTrans.AuthNetworkID"
    Alias ="Expr50"
    Expression ="pcChargeTrans.AuthSource"
    Alias ="Expr51"
    Expression ="pcChargeTrans.FMReferenceNum"
    Alias ="Expr52"
    Expression ="pcChargeTrans.CheckInDate"
    Alias ="Expr53"
    Expression ="pcChargeTrans.CheckOutDate"
    Alias ="Expr54"
    Expression ="pcChargeTrans.RoomRateAmount"
    Alias ="Expr55"
    Expression ="pcChargeTrans.CVV2_Resp"
    Alias ="Expr56"
    Expression ="pcChargeTrans.Selected"
    Alias ="Expr57"
    Expression ="pcChargeTrans.Estimated_Gratuity"
    Alias ="Expr58"
    Expression ="pcChargeTrans.Commercial_Card"
    Alias ="Expr59"
    Expression ="pcChargeTrans.Offline"
    Alias ="Expr60"
    Expression ="pcChargeTrans.Status"
    Alias ="Expr61"
    Expression ="pcChargeTrans.TroutD"
    Alias ="Expr62"
    Expression ="pcChargeTrans.Comment"
    Alias ="Expr63"
    Expression ="pcChargeTrans.CardPresent"
    Alias ="Expr64"
    Expression ="pcChargeTrans.CardholderID"
    Alias ="Expr65"
    Expression ="pcChargeTrans.AccountDataSrc"
    Alias ="Expr66"
    Expression ="pcChargeTrans.ECommIndicator"
    Alias ="Expr67"
    Expression ="pcChargeTrans.Business_Type"
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
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr31"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr32"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr33"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr34"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr35"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr36"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr37"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr38"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr39"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr40"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr41"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr42"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr43"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr44"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr45"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr46"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr47"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr48"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr49"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr50"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr51"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr52"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr53"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr54"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr55"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr56"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr57"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr58"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr59"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr61"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr62"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr63"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr64"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr65"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr66"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr67"
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
    Bottom =105
    Left =0
    Top =0
    ColumnsShown =539
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
