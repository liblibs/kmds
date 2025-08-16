Operation =1
Option =0
Where ="(((qryPurchaseOrderItemInfo.VendorID)=[parVendorID]) AND ((qryPurchaseOrderItemI"
    "nfo.Avail)<=[OrderAt]) AND ((qryPurchaseOrderItemInfo.OrderTo)>0) AND ((qryPurch"
    "aseOrderItemInfo.SOQ)>0))"
Begin InputTables
    Name ="qryPurchaseOrderItemInfo"
End
Begin OutputColumns
    Expression ="qryPurchaseOrderItemInfo.VendorID"
    Expression ="qryPurchaseOrderItemInfo.Name"
    Expression ="qryPurchaseOrderItemInfo.ItemID"
    Expression ="qryPurchaseOrderItemInfo.SalesDepartmentID"
    Expression ="qryPurchaseOrderItemInfo.Description"
    Expression ="qryPurchaseOrderItemInfo.SizeID"
    Expression ="qryPurchaseOrderItemInfo.Onhand"
    Expression ="qryPurchaseOrderItemInfo.Ordered"
    Expression ="qryPurchaseOrderItemInfo.Reserved"
    Expression ="qryPurchaseOrderItemInfo.Avail"
    Expression ="qryPurchaseOrderItemInfo.OrderAt"
    Expression ="qryPurchaseOrderItemInfo.OrderTo"
    Expression ="qryPurchaseOrderItemInfo.SOQRaw"
    Expression ="qryPurchaseOrderItemInfo.OrderMultiple"
    Expression ="qryPurchaseOrderItemInfo.SOQ"
    Expression ="qryPurchaseOrderItemInfo.UnitsPerCase"
    Expression ="qryPurchaseOrderItemInfo.SPrice"
    Expression ="qryPurchaseOrderItemInfo.SPriceStar"
    Expression ="qryPurchaseOrderItemInfo.ExtPrice"
    Expression ="qryPurchaseOrderItemInfo.Cost"
    Expression ="qryPurchaseOrderItemInfo.ExtCost"
    Alias ="Repack"
    Expression ="IIf((([soq]\\[UnitsPerCase])*[UnitsPerCase])=[soq],0,-1)"
End
Begin Parameters
    Name ="parVendorID"
    Flag =4
End
Begin OrderBy
    Expression ="qryPurchaseOrderItemInfo.SalesDepartmentID"
    Flag =0
    Expression ="qryPurchaseOrderItemInfo.Description"
    Flag =0
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
        dbText "Name" ="qryPurchaseOrderItemInfo.Onhand"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Ordered"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Reserved"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Avail"
        dbInteger "ColumnWidth" ="630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.OrderAt"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.OrderTo"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.SOQRaw"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.OrderMultiple"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.SOQ"
        dbInteger "ColumnWidth" ="585"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Name"
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.ItemID"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Description"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.SizeID"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.UnitsPerCase"
        dbInteger "ColumnWidth" ="465"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.SPrice"
        dbInteger "ColumnWidth" ="405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.SPriceStar"
        dbInteger "ColumnWidth" ="420"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.ExtPrice"
        dbInteger "ColumnWidth" ="720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.Cost"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPurchaseOrderItemInfo.ExtCost"
        dbInteger "ColumnWidth" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
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
    Bottom =173
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =192
        Bottom =180
        Top =0
        Name ="qryPurchaseOrderItemInfo"
        Name =""
    End
End
