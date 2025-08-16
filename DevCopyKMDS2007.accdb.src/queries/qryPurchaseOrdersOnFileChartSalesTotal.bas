Operation =1
Option =0
Where ="(((ItemMovement.Code)=Forms!PurchaseOrdersOnFile!PurchaseOrdersOnFileSub.Form!It"
    "emID) And ((ItemMovement.Description)=\"  SALES -- TOTAL\"))"
Begin InputTables
    Name ="ItemMovement"
    Name ="qryItemOnOrder"
    Name ="Item"
End
Begin OutputColumns
    Expression ="ItemMovement.Code"
    Expression ="ItemMovement.Description"
    Alias ="Ordered"
    Expression ="IIf(IsNull([OnOrder]),0,[OnOrder])"
    Expression ="Item.Onhand"
    Expression ="Item.OrderAt"
    Expression ="Item.OrderTo"
    Expression ="ItemMovement.CurrMM"
    Expression ="ItemMovement.Jan"
    Expression ="ItemMovement.Feb"
    Expression ="ItemMovement.Mar"
    Expression ="ItemMovement.Apr"
    Expression ="ItemMovement.May"
    Expression ="ItemMovement.Jun"
    Expression ="ItemMovement.Jul"
    Expression ="ItemMovement.Aug"
    Expression ="ItemMovement.Sep"
    Expression ="ItemMovement.Oct"
    Expression ="ItemMovement.Nov"
    Expression ="ItemMovement.Dec"
End
Begin Joins
    LeftTable ="ItemMovement"
    RightTable ="qryItemOnOrder"
    Expression ="ItemMovement.Code=qryItemOnOrder.ItemID"
    Flag =1
    LeftTable ="ItemMovement"
    RightTable ="Item"
    Expression ="ItemMovement.Code=Item.ID"
    Flag =1
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
        dbText "Name" ="Ordered"
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
    Bottom =174
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =232
        Bottom =180
        Top =0
        Name ="ItemMovement"
        Name =""
    End
    Begin
        Left =270
        Top =6
        Right =366
        Bottom =90
        Top =0
        Name ="qryItemOnOrder"
        Name =""
    End
    Begin
        Left =359
        Top =73
        Right =455
        Bottom =187
        Top =0
        Name ="Item"
        Name =""
    End
End
