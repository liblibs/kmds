Operation =1
Option =0
Where ="(((ItemSalePrice.StartDate)<Now()) AND ((ItemSalePrice.EndDate)>Now()))"
Begin InputTables
    Name ="ItemSalePrice"
End
Begin OutputColumns
    Expression ="ItemSalePrice.ItemID"
    Expression ="ItemSalePrice.StartDate"
    Expression ="ItemSalePrice.EndDate"
    Expression ="ItemSalePrice.SalePrice"
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
    Bottom =137
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =435
        Top =20
        Right =625
        Bottom =119
        Top =0
        Name ="ItemSalePrice"
        Name =""
    End
End
