Operation =3
Name ="BuyDownItem"
Option =0
Where ="((([ConvCIGSLINE].[Order])=[parOrder]) AND ((BuyDownItem.ItemsCode) Is Null))"
Begin InputTables
    Name ="ConvCIGSLINE"
    Name ="Item"
    Name ="BuyDownItem"
End
Begin OutputColumns
    Alias ="BuyDownID"
    Name ="BuyDownID"
    Expression ="[parBuyDownID]"
    Alias ="Expr1"
    Name ="Line"
    Expression ="ConvCIGSLINE.Line"
    Alias ="Expr2"
    Name ="Order"
    Expression ="ConvCIGSLINE.Order"
    Alias ="ItemsCode"
    Name ="ItemsCode"
    Expression ="ConvCIGSLINE.Item"
    Alias ="Expr3"
    Name ="QuantitySold"
    Expression ="ConvCIGSLINE.ShpQty"
    Alias ="Expr4"
    Name ="DiscountDollars"
    Expression ="ConvCIGSLINE.ShpDollars"
    Alias ="Expr5"
    Name ="SpcPrice"
    Expression ="ConvCIGSLINE.SpcPrice"
End
Begin Parameters
    Name ="parBuyDownID"
    Flag =3
End
Begin Joins
    LeftTable ="Item"
    RightTable ="BuyDownItem"
    Expression ="Item.ID = BuyDownItem.ItemsCode"
    Flag =2
End
Begin OrderBy
    Expression ="ConvCIGSLINE.Line"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="setupDate"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BuyDownID"
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
        dbText "Name" ="ItemsCode"
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
    Bottom =321
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =143
        Bottom =45
        Top =0
        Name ="ConvCIGSLINE"
        Name =""
    End
    Begin
        Left =181
        Top =6
        Right =277
        Bottom =165
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =315
        Top =6
        Right =411
        Bottom =255
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
End
