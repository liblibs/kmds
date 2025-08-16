Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11820
    DatasheetFontHeight =10
    ItemSuffix =54
    Left =2595
    Top =5010
    Right =15885
    Bottom =8175
    DatasheetBackColor =16777164
    OrderBy ="LineID"
    RecSrcDt = Begin
        0x7d71956fb56fe340
    End
    RecordSource ="SELECT OrderLine.ID, OrderLine.OrdheadID, OrderLine.LineID, OrderLine.ItemID, Or"
        "derLine.QuantityShipped, OrderLine.Price, OrderLine.SalePrice, OrderLine.Buydown"
        ", OrderLine.Cost AS OrderCost, [QuantityShipped]*(OrderLine.Price-[BuyDown]) AS "
        "ExtPrice, [QuantityShipped]*OrderLine.Cost AS ExtCost, OrderLine.QuantityOrdered"
        " FROM OrderLine; "
    Caption ="OrderLine"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    MenuBar ="KMDS"
    DatasheetFontName ="Trebuchet MS"
    OnKeyDown ="[Event Procedure]"
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777164
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
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
            Height =465
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =780
                    Height =240
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =900
                    Top =60
                    Width =480
                    Height =240
                    Name ="OrderID_Label"
                    Caption ="OrderID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1440
                    Top =60
                    Width =480
                    Height =240
                    Name ="LineID_Label"
                    Caption ="LineID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1980
                    Top =60
                    Width =480
                    Height =240
                    Name ="ItemID_Label"
                    Caption ="ItemID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3240
                    Top =60
                    Width =660
                    Height =240
                    Name ="QuantityShipped_Label"
                    Caption ="QuantityShipped"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3960
                    Top =60
                    Width =1260
                    Height =240
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5280
                    Top =60
                    Width =1260
                    Height =240
                    Name ="Cost_Label"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6600
                    Top =60
                    Width =480
                    Height =240
                    Name ="SetupDate_Label"
                    Caption ="SetupDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7140
                    Top =60
                    Width =720
                    Height =240
                    Name ="ChangeDate_Label"
                    Caption ="ChangeDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7920
                    Top =60
                    Width =1260
                    Height =240
                    Name ="Label21"
                    Caption ="ExtPrice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =9360
                    Top =60
                    Width =780
                    Height =240
                    Name ="Label49"
                    Caption ="Buydown:"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =300
                    Top =60
                    Width =1095
                    Height =405
                    Name ="cmdAddRecord"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =2220
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =585
                    ColumnOrder =0
                    BackColor =16777164
                    ForeColor =0
                    Name ="LineID"
                    ControlSource ="LineID"
                    StatusBarText ="Double click Line number to insert a line here."
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Double click Line number to insert a line here."
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    DatasheetCaption ="Id"
                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =600
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =720
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =16777164
                    ForeColor =0
                    Name ="Quantity"
                    ControlSource ="QuantityShipped"
                    StatusBarText ="Press {F8} to Search or Add New"
                    ValidationRule ="Is Not Null"
                    ValidationText ="Blank or Null is not allowed here; try using the number '0'."
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    DatasheetCaption ="Qty"
                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7500
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =8
                    TabIndex =2
                    BackColor =16777164
                    ForeColor =0
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8760
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =810
                    ColumnOrder =10
                    TabIndex =3
                    BackColor =16777164
                    ForeColor =0
                    Name ="Cost"
                    ControlSource ="OrderCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9660
                    Top =60
                    Width =960
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =11
                    TabIndex =4
                    BackColor =16777164
                    ForeColor =0
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9120
                            Top =540
                            Width =705
                            Height =210
                            Name ="Label45"
                            Caption ="Ext Price"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =10680
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =855
                    ColumnOrder =12
                    TabIndex =5
                    BackColor =16777164
                    ForeColor =0
                    Name ="ExtCost"
                    ControlSource ="ExtCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =15
                    ListRows =10
                    ListWidth =10800
                    Left =1380
                    Top =60
                    Width =1080
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =2
                    TabIndex =6
                    BoundColumn =12
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM lkupCluOrderItemByCode;"
                    ColumnWidths ="2160;4320;720;1080;1080;720;720;720;720;720;720;720;720;720;720"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1440
                            Top =420
                            Width =840
                            Height =240
                            Name ="Label46"
                            Caption ="Code #"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =6
                    ListRows =10
                    ListWidth =10800
                    Left =2520
                    Top =60
                    Width =3780
                    ColumnWidth =1410
                    ColumnOrder =3
                    TabIndex =7
                    BoundColumn =1
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="Description"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM lkupCluOrderItemByDesc; "
                    ColumnWidths ="4320;720;1080;720;720;720"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ColumnCount =8
                    ListRows =1
                    ListWidth =1800
                    Left =6360
                    Top =60
                    Width =1080
                    Height =255
                    ColumnWidth =1000
                    ColumnOrder =4
                    TabIndex =9
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="Size"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT * FROM lkupCluOrderItemByPkgSize; "
                    ColumnWidths ="0;1440"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8280
                    Top =60
                    Width =180
                    ColumnWidth =-2
                    ColumnOrder =7
                    TabIndex =10
                    BackColor =16777164
                    Name ="SalePrice"
                    ControlSource ="SalePrice"
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000002000000000000000400000001000000 ,
                        0x00000000ff000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x27002a00270000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ff000000030000002700 ,
                        0x2a002700000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7860
                            Top =360
                            Width =285
                            Height =240
                            Name ="Label43"
                            Caption ="SP"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =8
                    ListRows =25
                    ListWidth =1800
                    Left =6360
                    Top =435
                    Width =1080
                    Height =255
                    ColumnWidth =1320
                    ColumnOrder =6
                    TabIndex =11
                    BackColor =16777164
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="Category"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, IIf(IsNull(SubCategory.Description),\" Not Assi\",Left(SubCatego"
                        "ry.Description,8)) AS Descr FROM Item LEFT JOIN SubCategory ON Item.SubCategoryI"
                        "D=SubCategory.ID; "
                    ColumnWidths ="0;1440"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =480
                    ColumnWidth =-2
                    ColumnOrder =9
                    TabIndex =12
                    BackColor =16777164
                    ForeColor =0
                    Name ="Buydown"
                    ControlSource ="Buydown"
                    Format ="#,##0.00;($#,##0.00)[Red];\"\";\015\012"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x00000000ffff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =15
                    ListRows =1
                    Left =6120
                    Top =780
                    Width =720
                    ColumnWidth =-2
                    ColumnOrder =5
                    TabIndex =8
                    BoundColumn =12
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
                    RowSource ="SELECT * FROM lkupCluOrderItemByCode; "
                    ColumnWidths ="0;0;0;1080;0;0;0;0;0;0;0;0;0;0;0;0"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff0100000000000000ffff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6180
                            Top =1140
                            Width =840
                            Height =240
                            Name ="Label53"
                            Caption ="Shelf"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =630
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Width =780
                    Height =255
                    ColumnWidth =1440
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    OverlapFlags =85
                    Left =960
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="OrdheadID"
                    ControlSource ="OrdheadID"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    Left =8820
                    Width =1260
                    Height =255
                    TabIndex =2
                    Name ="txtSumExtPrice"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    Left =7500
                    Width =1260
                    Height =255
                    TabIndex =3
                    Name ="txtSumExtCost"
                    ControlSource ="=Sum([ExtCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =5340
                    Width =1980
                    Height =255
                    TabIndex =4
                    Name ="txtSumExtQuantity"
                    ControlSource ="=Sum([QuantityShipped])"
                    Format ="Standard"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3180
                    Top =180
                    Width =1095
                    Height =405
                    TabIndex =5
                    Name ="cmdAddNewRecord"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrderSub.cls"
