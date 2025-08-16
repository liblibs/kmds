Operation =1
Option =0
Where ="(((Item.Location) Between Forms!PrintInventoryReport!cboLocationBegin And Forms!"
    "PrintInventoryReport!cboLocationEnd) And ((Item.VendorID) Between Forms!PrintInv"
    "entoryReport!cboVendorIDBegin And Forms!PrintInventoryReport!cboVendorIDEnd))"
Begin InputTables
    Name ="Item"
    Name ="Size"
    Name ="qryItemMovementWithTransfersReversed"
End
Begin OutputColumns
    Expression ="Item.Location"
    Expression ="Item.ID"
    Expression ="Item.Description"
    Alias ="SizeD"
    Expression ="Size.Description"
    Alias ="OnTheShelf"
    Expression ="[onhand]+[TransferredToday]-[SoldTodayCr1]-[SoldTodayCr2]-[SoldTodayCr3]-[SoldTo"
        "dayCr4]"
    Expression ="Item.VendorID"
    Expression ="qryItemMovementWithTransfersReversed.CurrMMQty"
    Expression ="qryItemMovementWithTransfersReversed.JanQty"
    Expression ="qryItemMovementWithTransfersReversed.FebQty"
    Expression ="qryItemMovementWithTransfersReversed.MarQty"
    Expression ="qryItemMovementWithTransfersReversed.AprQty"
    Expression ="qryItemMovementWithTransfersReversed.MayQty"
    Expression ="qryItemMovementWithTransfersReversed.JunQty"
    Expression ="qryItemMovementWithTransfersReversed.JulQty"
    Expression ="qryItemMovementWithTransfersReversed.AugQty"
    Expression ="qryItemMovementWithTransfersReversed.SepQty"
    Expression ="qryItemMovementWithTransfersReversed.OctQty"
    Expression ="qryItemMovementWithTransfersReversed.NovQty"
    Expression ="qryItemMovementWithTransfersReversed.DecQty"
    Expression ="qryItemMovementWithTransfersReversed.DecQty"
    Expression ="qryItemMovementWithTransfersReversed.DecQty"
    Expression ="qryItemMovementWithTransfersReversed.YtdQ"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="qryItemMovementWithTransfersReversed"
    Expression ="Item.ID = qryItemMovementWithTransfersReversed.Code"
    Flag =1
End
Begin OrderBy
    Expression ="Item.Location"
    Flag =0
    Expression ="Item.Description"
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
        dbText "Name" ="OnTheShelf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeD"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-1
        Top =13
        Right =165
        Bottom =262
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =195
        Top =145
        Right =291
        Bottom =259
        Top =0
        Name ="Size"
        Name =""
    End
    Begin
        Left =269
        Top =7
        Right =564
        Bottom =121
        Top =0
        Name ="qryItemMovementWithTransfersReversed"
        Name =""
    End
End
