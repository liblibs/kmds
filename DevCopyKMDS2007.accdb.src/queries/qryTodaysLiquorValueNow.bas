Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
    Name ="qryItemsSalePriceNow"
End
Begin OutputColumns
    Alias ="CurrentLiquorValue"
    Expression ="Sum(IIf(IsNull(qryItemsSalePriceNow.saleprice),[Price],qryItemsSalePriceNow.Sale"
        "Price)*[onhand])"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="qryItemsSalePriceNow"
    Expression ="Item.ID = qryItemsSalePriceNow.ItemID"
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
        dbText "Name" ="CurrentLiquorValue"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =183
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =316
        Top =23
        Right =481
        Bottom =182
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =550
        Top =23
        Right =646
        Bottom =137
        Top =0
        Name ="qryItemsSalePriceNow"
        Name =""
    End
End
