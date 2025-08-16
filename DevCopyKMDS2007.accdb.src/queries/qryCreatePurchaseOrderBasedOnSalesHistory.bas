dbMemo "SQL" ="PARAMETERS parBeginDate DateTime, parEndDate DateTime, parVendorID Long, parBegi"
    "nSalesDepartmentID Long, parEndSalesDepartmentID Long;\015\012SELECT Item.Vendor"
    "ID, Item.ID AS ItemID, SUM(InvoLine.Quantity) AS QuantityOrdered, SUM(InvoLine.Q"
    "uantity) AS QuantityShipped, IIf(SUM(InvoLine.Quantity)<>0,SUM(InvoLine.Price)/S"
    "UM(InvoLine.Quantity),SUM(InvoLine.Price)) AS Price, IIf(SUM(InvoLine.Quantity)<"
    ">0,SUM(Item.Cost*InvoLine.Quantity)/SUM(InvoLine.Quantity),SUM(Item.Cost)) AS It"
    "emCost, IIf(SUM(InvoLine.Quantity) Mod SUM(Item.UnitsPerCase)<>0,True,False) AS "
    "Repack, 0 AS LineNumber INTO PurchaseOrderItemDaysSupply\015\012FROM (((InvoLine"
    " INNER JOIN TranTypes ON InvoLine.TransType=TranTypes.Code) INNER JOIN Item ON I"
    "nvoLine.Code=Item.ID) INNER JOIN SalesDepartment ON Item.SalesDepartmentID=Sales"
    "Department.ID) INNER JOIN SalesGroup ON SalesDepartment.SalesGroupID=SalesGroup."
    "ID\015\012WHERE (InvoLine.TransType='S' Or InvoLine.TransType='C') And (InvoLine"
    ".SetupDate Between [parBeginDate] And [parEndDate]) And (Item.SalesDepartmentID "
    "Between [parBeginSalesDepartmentID] And [parEndSalesDepartmentID]) And (SalesGro"
    "up.Description='Liquor')\015\012GROUP BY Item.VendorID, Item.ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "LogMessages" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityOrdered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="nLineCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineNumber"
        dbLong "AggregateType" ="-1"
    End
End
