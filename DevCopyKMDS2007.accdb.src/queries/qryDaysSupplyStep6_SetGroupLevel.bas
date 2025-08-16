Operation =4
Option =0
Where ="(((DaysSupplyItemSales.Sold) Between [parLowerAmount] And [parUpperAmount]))"
Begin InputTables
    Name ="Item"
    Name ="DaysSupplyItemSales"
End
Begin OutputColumns
    Name ="Item.DaysSupplyGroupID"
    Expression ="[parDaysSupplyGroupLevel]"
End
Begin Parameters
    Name ="parDaysSupplyGroupLevel"
    Flag =4
    Name ="parLowerAmount"
    Flag =4
    Name ="parUpperAmount"
    Flag =4
End
Begin Joins
    LeftTable ="Item"
    RightTable ="DaysSupplyItemSales"
    Expression ="Item.ID = DaysSupplyItemSales.Code"
    Flag =1
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
    Bottom =307
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =152
        Bottom =315
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =369
        Top =6
        Right =585
        Bottom =90
        Top =0
        Name ="DaysSupplyItemSales"
        Name =""
    End
End
