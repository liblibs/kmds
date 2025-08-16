Operation =4
Option =0
Where ="(((BuyDownItem.ItemsCode)=[parItemID]))"
Begin InputTables
    Name ="BuyDownItem"
End
Begin OutputColumns
    Name ="BuyDownItem.QuantitySold"
    Expression ="BuyDownItem.QuantitySold+[parQuantity]"
    Name ="BuyDownItem.DiscountDollars"
    Expression ="BuyDownItem.DiscountDollars+[parDollars]"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
    Name ="parQuantity"
    Flag =3
    Name ="parDollars"
    Flag =5
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
    Bottom =126
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =50
        Top =1
        Right =184
        Bottom =145
        Top =0
        Name ="BuyDownItem"
        Name =""
    End
End
