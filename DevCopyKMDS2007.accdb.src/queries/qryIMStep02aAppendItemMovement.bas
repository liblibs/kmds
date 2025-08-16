Operation =3
Name ="ItemMovement"
Option =0
Begin InputTables
    Name ="qryItemMovement"
    Name ="TranTypes"
End
Begin OutputColumns
    Name ="Code"
    Expression ="qryItemMovement.Code"
    Name ="Description"
    Expression ="TranTypes.Description"
    Alias ="CurrMM"
    Name ="CurrMM"
    Expression ="0"
    Alias ="YTDQty"
    Name ="YTDQty"
    Expression ="Sum(qryItemMovement.SumOfQuantity)"
    Alias ="Jan"
    Name ="Jan"
    Expression ="Sum(IIf([MM]=\"01\",[SumOfQuantity],0))"
    Alias ="Feb"
    Name ="Feb"
    Expression ="Sum(IIf([MM]=\"02\",[SumOfQuantity],0))"
    Alias ="Mar"
    Name ="Mar"
    Expression ="Sum(IIf([MM]=\"03\",[SumOfQuantity],0))"
    Alias ="Apr"
    Name ="Apr"
    Expression ="Sum(IIf([MM]=\"04\",[SumOfQuantity],0))"
    Alias ="May"
    Name ="May"
    Expression ="Sum(IIf([MM]=\"05\",[SumOfQuantity],0))"
    Alias ="Jun"
    Name ="Jun"
    Expression ="Sum(IIf([MM]=\"06\",[SumOfQuantity],0))"
    Alias ="Jul"
    Name ="Jul"
    Expression ="Sum(IIf([MM]=\"07\",[SumOfQuantity],0))"
    Alias ="Aug"
    Name ="Aug"
    Expression ="Sum(IIf([MM]=\"08\",[SumOfQuantity],0))"
    Alias ="Sep"
    Name ="Sep"
    Expression ="Sum(IIf([MM]=\"09\",[SumOfQuantity],0))"
    Alias ="Oct"
    Name ="Oct"
    Expression ="Sum(IIf([MM]=\"10\",[SumOfQuantity],0))"
    Alias ="Nov"
    Name ="Nov"
    Expression ="Sum(IIf([MM]=\"11\",[SumOfQuantity],0))"
    Alias ="Dec"
    Name ="Dec"
    Expression ="Sum(IIf([MM]=\"12\",[SumOfQuantity],0))"
End
Begin Joins
    LeftTable ="qryItemMovement"
    RightTable ="TranTypes"
    Expression ="qryItemMovement.TransType = TranTypes.Code"
    Flag =2
End
Begin OrderBy
    Expression ="TranTypes.Description"
    Flag =0
End
Begin Groups
    Expression ="qryItemMovement.Code"
    GroupLevel =0
    Expression ="TranTypes.Description"
    GroupLevel =0
    Expression ="0"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Jan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Feb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mar"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Apr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="May"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jun"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jul"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Aug"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Nov"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YTDQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrMM"
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
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =655
    Begin
        Left =164
        Top =8
        Right =344
        Bottom =152
        Top =0
        Name ="qryItemMovement"
        Name =""
    End
    Begin
        Left =37
        Top =23
        Right =133
        Bottom =107
        Top =0
        Name ="TranTypes"
        Name =""
    End
End
