Operation =1
Option =0
Begin InputTables
    Name ="qryClubOrderGuidesReportAndSysfile"
    Name ="qryClubSalesThisMonthByItemByWeek"
End
Begin OutputColumns
    Expression ="qryClubOrderGuidesReportAndSysfile.*"
    Alias ="W1"
    Expression ="IIf(IsNull([Week1]),0,[Week1])"
    Alias ="W2"
    Expression ="IIf(IsNull([Week2]),0,[Week2])"
    Alias ="W3"
    Expression ="IIf(IsNull([Week3]),0,[Week3])"
    Alias ="W4"
    Expression ="IIf(IsNull([Week4]),0,[Week4])"
    Alias ="W5"
    Expression ="IIf(IsNull([Week5]),0,[Week5])"
    Alias ="STRIPPED"
    Expression ="IIf(([ItemID]>99900000000),[ItemID]-99900000000,[ItemID])"
End
Begin Joins
    LeftTable ="qryClubOrderGuidesReportAndSysfile"
    RightTable ="qryClubSalesThisMonthByItemByWeek"
    Expression ="qryClubOrderGuidesReportAndSysfile.ID = qryClubSalesThisMonthByItemByWeek.Club.I"
        "D"
    Flag =2
    LeftTable ="qryClubOrderGuidesReportAndSysfile"
    RightTable ="qryClubSalesThisMonthByItemByWeek"
    Expression ="qryClubOrderGuidesReportAndSysfile.ItemID = qryClubSalesThisMonthByItemByWeek.It"
        "em.ID"
    Flag =2
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
        dbText "Name" ="W1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="STRIPPED"
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
    Right =1193
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =242
        Bottom =120
        Top =0
        Name ="qryClubOrderGuidesReportAndSysfile"
        Name =""
    End
    Begin
        Left =311
        Top =0
        Right =407
        Bottom =114
        Top =0
        Name ="qryClubSalesThisMonthByItemByWeek"
        Name =""
    End
End
