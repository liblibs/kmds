Operation =1
Option =0
Begin InputTables
    Name ="LotteryScratchRecon"
    Name ="Item"
End
Begin OutputColumns
    Expression ="LotteryScratchRecon.BusinessDate"
    Expression ="LotteryScratchRecon.GameNumber"
    Expression ="LotteryScratchRecon.ItemID"
    Expression ="LotteryScratchRecon.BeginMeter"
    Expression ="LotteryScratchRecon.EndMeter"
    Expression ="LotteryScratchRecon.AddMeter"
    Expression ="LotteryScratchRecon.AmountSold"
    Expression ="LotteryScratchRecon.SetupDate"
    Expression ="LotteryScratchRecon.ChangeDate"
    Alias ="PhysicalAmount"
    Expression ="[BeginMeter]+[AddMeter]-[EndMeter]"
    Alias ="Variance"
    Expression ="[AmountSold]-([BeginMeter]+[AddMeter]-[EndMeter])"
    Alias ="VarianceDollars"
    Expression ="([AmountSold]-([BeginMeter]+[AddMeter]-[EndMeter]))*[Price]"
    Expression ="Item.Price"
    Expression ="Item.Description"
    Alias ="SoldDollars"
    Expression ="[amountsold]*[price]"
End
Begin Joins
    LeftTable ="LotteryScratchRecon"
    RightTable ="Item"
    Expression ="LotteryScratchRecon.ItemID=Item.ID"
    Flag =1
End
Begin OrderBy
    Expression ="LotteryScratchRecon.GameNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="PhysicalAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Variance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VarianceDollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldDollars"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="LotteryScratchRecon"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
