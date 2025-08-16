Operation =1
Option =0
Begin InputTables
    Name ="qryTodaysLiquorTransactions"
    Name ="qryTodaysLiquorValueNow"
End
Begin OutputColumns
    Expression ="qryTodaysLiquorTransactions.Today"
    Alias ="LastNight"
    Expression ="[CurrentLiquorValue]-[Counter]-[Clubs]-[Purchases]-[TransfersIn]-[TransfersOut]-"
        "[Adjustments]"
    Expression ="qryTodaysLiquorTransactions.Counter"
    Expression ="qryTodaysLiquorTransactions.Clubs"
    Expression ="qryTodaysLiquorTransactions.Purchases"
    Expression ="qryTodaysLiquorTransactions.TransfersOUT"
    Expression ="qryTodaysLiquorTransactions.TransfersIN"
    Expression ="qryTodaysLiquorTransactions.Adjustments"
    Expression ="qryTodaysLiquorValueNow.CurrentLiquorValue"
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
        dbText "Name" ="qryTodaysLiquorValueNow.CurrentLiquorValue"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryTodaysLiquorTransactions.Purchases"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryTodaysLiquorTransactions.Clubs"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryTodaysLiquorTransactions.Adjustments"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastNight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryTodaysLiquorTransactions.Today"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryTodaysLiquorTransactions.Counter"
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
    Bottom =78
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="qryTodaysLiquorTransactions"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =312
        Bottom =75
        Top =0
        Name ="qryTodaysLiquorValueNow"
        Name =""
    End
End
