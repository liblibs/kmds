Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15600
    RowHeight =270
    DatasheetFontHeight =10
    ItemSuffix =68
    Left =5250
    Top =5445
    Right =18405
    Bottom =7335
    DatasheetBackColor =13434828
    Filter ="QuantityShipped<>0"
    ShortcutMenuBar ="KMDSPopup"
    OrderBy ="ItemID"
    RecSrcDt = Begin
        0x79cedb1eb06be340
    End
    RecordSource ="SELECT PurchaseOrderLine.ID, PurchaseOrderLine.PurchaseOrderID, PurchaseOrderLin"
        "e.VendorID, PurchaseOrderLine.LineID, PurchaseOrderLine.ItemID, PurchaseOrderLin"
        "e.QuantityShipped, PurchaseOrderLine.Price, PurchaseOrderLine.Cost, PurchaseOrde"
        "rLine.Repack, PurchaseOrderLine.SetupDate, PurchaseOrderLine.ChangeDate, [Quanti"
        "tyShipped]*[Price] AS ExtPrice, [QuantityShipped]*[Cost] AS ExtCost, IIf([Price]"
        ">0,([Price]-[Cost])/[Price],0) AS GP, IIf([repack],[QuantityShipped]*[Price],0) "
        "AS ExtPriceRepack, IIf([repack],[QuantityShipped],0) AS QtyRepack, IIf([repack],"
        "0,[QuantityShipped]*[Price]) AS ExtPriceCase, IIf([repack],0,[QuantityShipped]) "
        "AS QtyCase, PurchaseOrderLine.PrintLabels FROM PurchaseOrderLine; "
    Caption ="PurchaseOrderLine subform"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnDblClick ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =13434828
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Line
            SpecialEffect =3
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionGroup
            SpecialEffect =3
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            SpecialEffect =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            BorderLineStyle =0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =765
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =180
                    Width =540
                    Height =240
                    Name ="ID_Label"
                    Caption ="ID"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =5820
                    Width =600
                    Height =240
                    Name ="QuantityOrdered_Label"
                    Caption ="QuantityOrdered"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =3
                    Left =6420
                    Width =1140
                    Height =240
                    Name ="Price_Label"
                    Caption ="Price"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =4800
                    Width =720
                    Height =255
                    ColumnOrder =0
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2940
                    Width =780
                    Height =255
                    ColumnOrder =1
                    TabIndex =1
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3960
                    Width =660
                    Height =255
                    ColumnOrder =2
                    TabIndex =2
                    Name ="PurchaseOrderID"
                    ControlSource ="PurchaseOrderID"
                    FontName ="Trebuchet MS"

                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =720
                    Width =420
                    Height =285
                    Name ="Label31"
                    Caption ="Item"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =3
                    Left =7560
                    Width =1200
                    Height =240
                    Name ="Label32"
                    Caption ="Price"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =8880
                    Width =1140
                    Height =240
                    Name ="Label33"
                    Caption ="Cost"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =3
                    Left =10020
                    Width =1200
                    Height =240
                    Name ="Label34"
                    Caption ="Price"
                    FontName ="Trebuchet MS"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =12540
                    Width =720
                    Height =240
                    Name ="Label41"
                    Caption ="Repack"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11460
                    Top =120
                    Width =930
                    Height =405
                    TabIndex =3
                    Name ="cmdOpenItem"
                    Caption ="OpenItem"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =10080
                    Top =240
                    Width =1095
                    Height =405
                    TabIndex =4
                    Name ="cmdAddNew"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =120
                    Width =795
                    Height =405
                    TabIndex =5
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"

                End
            End
        End
        Begin Section
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Width =360
                    Height =255
                    ColumnWidth =465
                    ColumnOrder =0
                    BackColor =13434828
                    Name ="LineID"
                    ControlSource ="LineID"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =420
                    Width =480
                    Height =255
                    ColumnWidth =690
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    Name ="QuantityShipped"
                    ControlSource ="QuantityShipped"
                    StatusBarText ="Press {F8} to Search or Add New; {F9} to Filter 0 quantity; {F10} Extended Cost;"
                        " {F11} Extended Price"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Blank or Null is not allowed here; try using the number '0'."
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"
                    ControlTipText ="Double click to filter on 0 quantity"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =300
                            Width =660
                            Height =210
                            Name ="Label42"
                            Caption ="Qty"
                        End
                    End
                End
                Begin TextBox
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =8280
                    Width =660
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =6
                    TabIndex =6
                    BackColor =13434828
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =9540
                    Width =1140
                    Height =255
                    ColumnWidth =1000
                    ColumnOrder =8
                    FontWeight =700
                    TabIndex =8
                    BackColor =16776960
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =10
                    ListRows =10
                    ListWidth =10800
                    Left =960
                    Width =1080
                    Height =255
                    ColumnWidth =1000
                    ColumnOrder =2
                    TabIndex =2
                    BoundColumn =8
                    BackColor =13434828
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByCode.CodeOrUPC, lkupCluOrderItemByCode.Description, lku"
                        "pCluOrderItemByCode.OnShelf, lkupCluOrderItemByCode.Cost, lkupCluOrderItemByCode"
                        ".CurPrice, lkupCluOrderItemByCode.Star, lkupCluOrderItemByCode.StockNumber, lkup"
                        "CluOrderItemByCode.UnitsPerCase, lkupCluOrderItemByCode.ID, lkupCluOrderItemByCo"
                        "de.ItemVendorID FROM lkupCluOrderItemByCode; "
                    ColumnWidths ="1440;4320;720;720;720;720;1800;720;720;720"
                    StatusBarText ="Press {F8} to Search or Add New  ---  Double click to go to Item Setup/Change fo"
                        "rm"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    OnNotInList ="[Event Procedure]"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =9
                    ListRows =10
                    ListWidth =10800
                    Left =2100
                    Width =4140
                    Height =255
                    ColumnWidth =3465
                    ColumnOrder =3
                    TabIndex =3
                    BoundColumn =1
                    BackColor =13434828
                    Name ="ItemDescription"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByDesc.Description, lkupCluOrderItemByDesc.ID, lkupCluOrd"
                        "erItemByDesc.OnShelf, lkupCluOrderItemByDesc.Cost, lkupCluOrderItemByDesc.CurPri"
                        "ce, lkupCluOrderItemByDesc.Star, 1 AS Stocknumber, 1 AS UnitsPerCase, lkupCluOrd"
                        "erItemByDesc.ItemVendorID FROM lkupCluOrderItemByDesc ORDER BY lkupCluOrderItemB"
                        "yDesc.Description; "
                    ColumnWidths ="4320;1440;1440;1440;1440;1440;1440;1440"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =10740
                    Width =1140
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =9
                    TabIndex =9
                    BackColor =13434828
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="Standard"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =11940
                    Width =1140
                    Height =255
                    ColumnWidth =1000
                    ColumnOrder =10
                    TabIndex =10
                    BackColor =13434828
                    Name ="ExtCost"
                    ControlSource ="ExtCost"
                    Format ="Standard"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    AutoTab = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =13140
                    Width =1140
                    Height =255
                    ColumnWidth =840
                    ColumnOrder =11
                    TabIndex =11
                    BackColor =13434828
                    Name ="GP"
                    ControlSource ="GP"
                    Format ="Percent"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin CheckBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =14460
                    Top =60
                    ColumnWidth =840
                    ColumnOrder =12
                    TabIndex =12
                    Name ="Repack"
                    ControlSource ="Repack"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14160
                    Top =300
                    ColumnOrder =15
                    TabIndex =13
                    Name ="ExtPriceRepack"
                    ControlSource ="ExtPriceRepack"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12720
                            Top =300
                            Width =1275
                            Height =240
                            Name ="Label43"
                            Caption ="ExtPriceRepack:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14160
                    Top =600
                    ColumnOrder =16
                    TabIndex =14
                    Name ="QtyRepack"
                    ControlSource ="QtyRepack"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12720
                            Top =600
                            Width =930
                            Height =240
                            Name ="Label44"
                            Caption ="QtyRepack:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14160
                    Top =900
                    ColumnOrder =17
                    TabIndex =15
                    Name ="ExtPriceCase"
                    ControlSource ="ExtPriceCase"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12720
                            Top =900
                            Width =1065
                            Height =240
                            Name ="Label45"
                            Caption ="ExtPriceCase:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14160
                    Top =1200
                    ColumnOrder =18
                    TabIndex =16
                    Name ="QtyCase"
                    ControlSource ="QtyCase"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12720
                            Top =1200
                            Width =720
                            Height =240
                            Name ="Label46"
                            Caption ="QtyCase:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =8
                    ListRows =25
                    ListWidth =10080
                    Left =6300
                    Width =720
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =4
                    TabIndex =4
                    BoundColumn =1
                    BackColor =13434828
                    Name ="Size"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByPkgSize.PkgSize, lkupCluOrderItemByPkgSize.ID FROM lkup"
                        "CluOrderItemByPkgSize; "
                    ColumnWidths ="1440;1440;1440;1440;1440;1440;1440;1440"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ColumnCount =8
                    ListRows =25
                    ListWidth =720
                    Left =9000
                    Width =480
                    Height =255
                    ColumnWidth =500
                    ColumnOrder =7
                    TabIndex =7
                    BackColor =255
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000003000000000000000400000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x27002a00270000000000
                    End
                    Name ="Sl"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryPurchaseOrderItemsWithTodaysPrice.ID, qryPurchaseOrderItemsWithTodaysP"
                        "rice.Star FROM qryPurchaseOrderItemsWithTodaysPrice WHERE (((qryPurchaseOrderIte"
                        "msWithTodaysPrice.Star)>\"\")) ORDER BY qryPurchaseOrderItemsWithTodaysPrice.ID;"
                        " "
                    ColumnWidths ="0;360"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000ccffcc00030000002700 ,
                        0x2a002700000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8820
                    Top =360
                    ColumnWidth =840
                    ColumnOrder =13
                    TabIndex =17
                    Name ="PrintLabels"
                    ControlSource ="PrintLabels"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9050
                            Top =330
                            Width =720
                            Height =240
                            Name ="Label59"
                            Caption ="Labels"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =15
                    ListRows =1
                    Left =7200
                    Width =720
                    ColumnWidth =600
                    ColumnOrder =5
                    TabIndex =5
                    BoundColumn =1
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cmbShelf"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByDesc.OnShelf, lkupCluOrderItemByDesc.ID FROM lkupCluOrd"
                        "erItemByDesc; "
                    ColumnWidths ="1080;0"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7260
                            Top =360
                            Width =840
                            Height =240
                            Name ="Label53"
                            Caption ="Shelf"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =162
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =15
                    ListRows =1
                    Left =1440
                    Top =566
                    Width =420
                    ColumnWidth =600
                    ColumnOrder =14
                    TabIndex =18
                    BoundColumn =1
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x0100000070000000020000000000000005000000000000000300000001000000 ,
                        0x00000000fff20000000000000500000004000000070000000100000000000000 ,
                        0xed1c240000000000000000000000000000000000000000000000000000000000 ,
                        0x32003000000000003100300000000000
                    End
                    Name ="cmbDaysSupply"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CInt([CurDaysSupply]) AS CDS, lkupCluOrderItemByCurrentDaysSupply.ID, lku"
                        "pCluOrderItemByCurrentDaysSupply.CurDaysSupply FROM lkupCluOrderItemByCurrentDay"
                        "sSupply; "
                    ColumnWidths ="1080;0"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =566
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =806
                    DatasheetCaption ="Days"
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000fff20000020000003200 ,
                        0x3000000000000000000000000000000000000000000000000000000500000001 ,
                        0x00000000000000ed1c2400020000003100300000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1500
                            Top =930
                            Width =960
                            Height =240
                            Name ="Label65"
                            Caption ="Days Supply"
                            LayoutCachedLeft =1500
                            LayoutCachedTop =930
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =1170
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =162
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =15
                    ListRows =1
                    Left =2716
                    Top =519
                    Width =990
                    TabIndex =19
                    BoundColumn =1
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x0100000070000000020000000000000005000000000000000300000001000000 ,
                        0x00000000fff20000000000000500000004000000070000000100000000000000 ,
                        0xed1c240000000000000000000000000000000000000000000000000000000000 ,
                        0x32003000000000003100300000000000
                    End
                    Name ="cmbStockNumber"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByStockNumber.StockNumber, lkupCluOrderItemByStockNumber."
                        "ID FROM lkupCluOrderItemByStockNumber; "
                    ColumnWidths ="1080;0"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    LayoutCachedLeft =2716
                    LayoutCachedTop =519
                    LayoutCachedWidth =3706
                    LayoutCachedHeight =759
                    DatasheetCaption ="Stock #"
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000fff20000020000003200 ,
                        0x3000000000000000000000000000000000000000000000000000000500000001 ,
                        0x00000000000000ed1c2400020000003100300000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2776
                            Top =883
                            Width =960
                            Height =240
                            Name ="Label67"
                            Caption ="Stock #"
                            LayoutCachedLeft =2776
                            LayoutCachedTop =883
                            LayoutCachedWidth =3736
                            LayoutCachedHeight =1123
                        End
                    End
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1320
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =5760
                    Top =60
                    Width =540
                    Height =255
                    Name ="txtSumQuantity"
                    ControlSource ="=Sum([QuantityShipped])"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =7320
                    Top =60
                    Width =1380
                    TabIndex =1
                    Name ="txtSumPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    Left =9720
                    Top =60
                    Width =1380
                    TabIndex =2
                    Name ="txtSumCost"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =180
                    TabIndex =3
                    Name ="txtSumRepackDollars"
                    ControlSource ="=Sum([ExtPriceRepack])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11580
                            Top =180
                            Width =615
                            Height =240
                            Name ="Label48"
                            Caption ="Text47:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =480
                    TabIndex =4
                    Name ="txtSumRepackQty"
                    ControlSource ="=Sum([QtyRepack])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11580
                            Top =480
                            Width =615
                            Height =240
                            Name ="Label50"
                            Caption ="Text49:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =780
                    TabIndex =5
                    Name ="txtSumCaseDollars"
                    ControlSource ="=Sum([ExtPriceCase])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11580
                            Top =780
                            Width =615
                            Height =240
                            Name ="Label52"
                            Caption ="Text51:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13020
                    Top =1080
                    TabIndex =6
                    Name ="txtSumCaseQty"
                    ControlSource ="=Sum([QtyCase])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11580
                            Top =1080
                            Width =615
                            Height =240
                            Name ="Label54"
                            Caption ="Text53:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrdersOnFileSub.cls"
