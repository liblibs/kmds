Operation =4
Option =0
Begin InputTables
    Name ="PSCInput"
    Name ="vwUPCItemsWithCurrentPricing"
End
Begin OutputColumns
    Name ="PSCInput.ItemID"
    Expression ="vwUPCItemsWithCurrentPricing.id"
    Name ="PSCInput.Description"
    Expression ="vwUPCItemsWithCurrentPricing.description"
    Name ="PSCInput.OnShelf"
    Expression ="vwUPCItemsWithCurrentPricing.onshelf"
    Name ="PSCInput.Price"
    Expression ="vwUPCItemsWithCurrentPricing.curprice"
    Name ="PSCInput.Cost"
    Expression ="vwUPCItemsWithCurrentPricing.cost"
    Name ="PSCInput.SalesDepartmentID"
    Expression ="vwUPCItemsWithCurrentPricing.SalesDepartmentID"
    Name ="PSCInput.Locked"
    Expression ="True"
End
Begin Joins
    LeftTable ="PSCInput"
    RightTable ="vwUPCItemsWithCurrentPricing"
    Expression ="PSCInput.Code = vwUPCItemsWithCurrentPricing.UPCCode"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="PSCInput.Code"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Descr"
        dbInteger "ColumnWidth" ="3330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PSCInput.Description"
        dbInteger "ColumnWidth" ="3330"
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
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =24
        Top =17
        Right =120
        Bottom =191
        Top =0
        Name ="PSCInput"
        Name =""
    End
    Begin
        Left =178
        Top =7
        Right =531
        Bottom =196
        Top =0
        Name ="vwUPCItemsWithCurrentPricing"
        Name =""
    End
End
