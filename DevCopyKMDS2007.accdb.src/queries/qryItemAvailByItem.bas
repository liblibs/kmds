Operation =1
Option =0
Where ="(((vwItemsWithCurrentPricing.ID)=[parItemID]))"
Begin InputTables
    Name ="vwItemsWithCurrentPricing"
End
Begin OutputColumns
    Expression ="vwItemsWithCurrentPricing.ID"
    Expression ="vwItemsWithCurrentPricing.Description"
    Expression ="vwItemsWithCurrentPricing.OnShelf"
    Expression ="vwItemsWithCurrentPricing.CurPrice"
    Expression ="vwItemsWithCurrentPricing.BuyDown"
    Expression ="vwItemsWithCurrentPricing.SalesDepartmentID"
    Expression ="vwItemsWithCurrentPricing.Price"
    Expression ="vwItemsWithCurrentPricing.Reserved"
End
Begin Parameters
    Name ="parItemID"
    Flag =4
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =194
        Bottom =120
        Top =0
        Name ="vwItemsWithCurrentPricing"
        Name =""
    End
End
