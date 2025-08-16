Operation =3
Name ="ItemMovement"
Option =0
Where ="(((Format([mmyy],\"yyyymm\"))>=Format((Now()-365),\"yyyymm\") And (Format([mmyy]"
    ",\"yyyymm\"))<Format((Now()),\"yyyymm\")))"
Begin InputTables
    Name ="qryItemMovementPackSales"
End
Begin OutputColumns
    Alias ="Desc"
    Name ="Description"
    Expression ="\" Sales - Packs\""
    Name ="Code"
    Expression ="qryItemMovementPackSales.CartonID"
    Alias ="YTDQty"
    Name ="YTDQty"
    Expression ="Sum(CInt(([PacksSold]+0.01)/[PacksPerCarton]))"
    Alias ="CurrMM"
    Name ="CurrMM"
    Expression ="0"
    Alias ="Jan"
    Name ="Jan"
    Expression ="Sum(IIf([MM]=\"01\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Feb"
    Name ="Feb"
    Expression ="Sum(IIf([MM]=\"02\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Mar"
    Name ="Mar"
    Expression ="Sum(IIf([MM]=\"03\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Apr"
    Name ="Apr"
    Expression ="Sum(IIf([MM]=\"04\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="May"
    Name ="May"
    Expression ="Sum(IIf([MM]=\"05\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Jun"
    Name ="Jun"
    Expression ="Sum(IIf([MM]=\"06\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Jul"
    Name ="Jul"
    Expression ="Sum(IIf([MM]=\"07\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Aug"
    Name ="Aug"
    Expression ="Sum(IIf([MM]=\"08\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Sep"
    Name ="Sep"
    Expression ="Sum(IIf([MM]=\"09\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Oct"
    Name ="Oct"
    Expression ="Sum(IIf([MM]=\"10\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Nov"
    Name ="Nov"
    Expression ="Sum(IIf([MM]=\"11\",[PacksSold]/[PacksPerCarton],0))"
    Alias ="Dec"
    Name ="Dec"
    Expression ="Sum(IIf([MM]=\"12\",[PacksSold]/[PacksPerCarton],0))"
End
Begin Groups
    Expression ="\" Sales - Packs\""
    GroupLevel =0
    Expression ="qryItemMovementPackSales.CartonID"
    GroupLevel =0
    Expression ="0"
    GroupLevel =0
    Expression ="qryItemMovementPackSales.PackID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryItemPackSalesWithCartonID.CartonID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemPackSalesWithCartonID.PackID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemPackSalesWithCartonID.Description"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemMovementPackSales.qryItemPackSalesWithCartonID.PackID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemMovementPackSales.PackID"
        dbInteger "ColumnOrder" ="4"
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
    ColumnsShown =655
    Begin
        Left =52
        Top =10
        Right =148
        Bottom =124
        Top =0
        Name ="qryItemMovementPackSales"
        Name =""
    End
End
