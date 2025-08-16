Operation =6
Option =0
Begin InputTables
    Name ="ClubHistoryByBottleByMonth"
End
Begin OutputColumns
    Expression ="ClubHistoryByBottleByMonth.ClubID"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.ClubName"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.CatDesc"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.Code"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.Description"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.SizeD"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.[Year-Month]"
    GroupLevel =1
    Alias ="SumOfQty"
    Expression ="Sum(ClubHistoryByBottleByMonth.Qty)"
    Alias ="TotalQty"
    Expression ="Sum(ClubHistoryByBottleByMonth.Qty)"
    GroupLevel =2
    Alias ="TotalDollars"
    Expression ="Sum(ClubHistoryByBottleByMonth.Dollars)"
    GroupLevel =2
End
Begin OrderBy
    Expression ="ClubHistoryByBottleByMonth.ClubName"
    Flag =0
    Expression ="ClubHistoryByBottleByMonth.CatDesc"
    Flag =0
    Expression ="ClubHistoryByBottleByMonth.Code"
    Flag =0
End
Begin Groups
    Expression ="ClubHistoryByBottleByMonth.ClubID"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.ClubName"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.CatDesc"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.Code"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.Description"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.SizeD"
    GroupLevel =2
    Expression ="ClubHistoryByBottleByMonth.[Year-Month]"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="ClubHistoryByBottleByMonth.ClubName"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubHistoryByBottleByMonth.Descr"
        dbInteger "ColumnWidth" ="3690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="01"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="02"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="03"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="04"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="05"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="06"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="07"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="08"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="09"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalDollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
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
    ColumnsShown =559
    Begin
        Left =38
        Top =6
        Right =313
        Bottom =120
        Top =0
        Name ="ClubHistoryByBottleByMonth"
        Name =""
    End
End
