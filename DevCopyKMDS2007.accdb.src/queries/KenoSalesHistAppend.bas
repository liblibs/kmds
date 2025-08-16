dbMemo "SQL" ="INSERT INTO InvoLine ( SetupDate, Quantity, InvoHeadID, ID, CashSysID, CashierID"
    ", ShiftID, Code, SalesDept, Price, TransType )\015\012SELECT \"9/2/2011\" AS Set"
    "upDate, KenoSalesHist.sep AS Quantity, 9999 AS InvoHeadID, KenoSalesHist.ID, 99 "
    "AS CashSysID, 99 AS CashierID, 99 AS ShiftID, KenoSalesHist.itemid AS Code, 1 AS"
    " SalesDept, [Item].[price]*[Quantity] AS Prices, \"S\" AS TransType\015\012FROM "
    "KenoSalesHist INNER JOIN Item ON KenoSalesHist.ItemID = Item.ID\015\012WHERE ((("
    "KenoSalesHist.sep)<>0));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="UPCItem.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.Description"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoHeadID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.ShiftID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShiftID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Prices"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.Uorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KenoSalesHist.nov"
        dbLong "AggregateType" ="-1"
    End
End
