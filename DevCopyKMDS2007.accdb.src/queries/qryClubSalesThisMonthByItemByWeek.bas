Operation =1
Option =0
Begin InputTables
    Name ="qryInvoLineClubSalesByItemByWeek"
    Name ="Club"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Club.ID"
    Expression ="Club.Name"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SumOfSumOfQuantity"
    Expression ="Sum(qryInvoLineClubSalesByItemByWeek.SumOfQuantity)"
    Alias ="Week1"
    Expression ="Sum(IIf([week]=1,[SumOfQuantity],0))"
    Alias ="Week2"
    Expression ="Sum(IIf([week]=2,[SumOfQuantity],0))"
    Alias ="Week3"
    Expression ="Sum(IIf([week]=3,[SumOfQuantity],0))"
    Alias ="Week4"
    Expression ="Sum(IIf([week]=4,[SumOfQuantity],0))"
    Alias ="Week5"
    Expression ="Sum(IIf([week]>4,[SumOfQuantity],0))"
End
Begin Joins
    LeftTable ="qryInvoLineClubSalesByItemByWeek"
    RightTable ="Club"
    Expression ="qryInvoLineClubSalesByItemByWeek.ClubId = Club.ID"
    Flag =1
    LeftTable ="qryInvoLineClubSalesByItemByWeek"
    RightTable ="Item"
    Expression ="qryInvoLineClubSalesByItemByWeek.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="Club.ID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="Item.ID"
    GroupLevel =0
    Expression ="Item.Description"
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
        dbText "Name" ="SumOfSumOfQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Week5"
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
    Bottom =283
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =224
        Top =0
        Name ="qryInvoLineClubSalesByItemByWeek"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =162
        Top =131
        Right =258
        Bottom =245
        Top =0
        Name ="Item"
        Name =""
    End
End
