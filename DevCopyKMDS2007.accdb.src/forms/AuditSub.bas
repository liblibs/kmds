Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13020
    DatasheetFontHeight =10
    ItemSuffix =43
    Left =11964
    Top =6708
    Right =26220
    Bottom =11808
    Filter ="([qryAudit].[OnShelf]<=0)"
    OrderBy ="[qryAudit].[PhysicalCount], [qryAudit].[ExtPrice]"
    RecSrcDt = Begin
        0x73d9ade35581e340
    End
    RecordSource ="qryAudit"
    Caption ="AuditSub"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
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
            Height =1200
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =3420
                    Top =300
                    Width =1560
                    Height =255
                    Name ="OnShelf_Label"
                    Caption ="OnShelf"
                End
            End
        End
        Begin Section
            Height =1620
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5280
                    Top =1140
                    Width =1140
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =15
                    Name ="ItemID"
                    ControlSource ="ItemID"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3660
                            Top =1140
                            Width =1560
                            Height =255
                            Name ="ItemID_Label"
                            Caption ="ItemID"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5820
                    Top =180
                    Width =900
                    Height =255
                    ColumnWidth =1590
                    ColumnOrder =3
                    TabIndex =2
                    BackColor =16777215
                    ForeColor =0
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6840
                    Top =180
                    Width =1260
                    Height =255
                    ColumnWidth =1044
                    ColumnOrder =8
                    TabIndex =3
                    BackColor =16777215
                    ForeColor =0
                    Name ="PhysicalCount"
                    ControlSource ="PhysicalCount"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =870
                            Width =870
                            Height =210
                            Name ="Label42"
                            Caption ="COUNT"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8220
                    Top =180
                    Width =720
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =9
                    TabIndex =4
                    BackColor =16777215
                    ForeColor =0
                    Name ="VarianceAmount"
                    ControlSource ="VarianceAmount"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000700000000300000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000500000002000000040000000100000000000000 ,
                        0xff99cc00000000000400000005000000070000000100000000000000ccffcc00 ,
                        0x00000000300000000000300000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000300000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000050000000100000000 ,
                        0x000000ff99cc0001000000300000000000000000000000000000000000000000 ,
                        0x000000000000040000000100000000000000ccffcc0001000000300000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8280
                            Top =480
                            Width =705
                            Height =210
                            Name ="Label31"
                            Caption ="Variance"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9120
                    Top =180
                    Width =1035
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =10
                    TabIndex =5
                    BackColor =16777215
                    ForeColor =0
                    Name ="Price"
                    ControlSource ="Price"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =180
                    Width =1035
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =12
                    TabIndex =6
                    BackColor =16777215
                    ForeColor =0
                    Name ="Cost"
                    ControlSource ="Cost"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2640
                    Top =1260
                    Width =1035
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =17
                    TabIndex =7
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1020
                            Top =1260
                            Width =1560
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11460
                    Top =180
                    Width =1035
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =14
                    TabIndex =8
                    BackColor =16777215
                    ForeColor =0
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="Short Time"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    ColumnCount =6
                    ListRows =25
                    ListWidth =10080
                    Left =120
                    Top =120
                    Width =1080
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =0
                    TabIndex =1
                    BoundColumn =5
                    BackColor =16777215
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cmbItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT vwUPCItemsWithCurrentPricing.UPCCode, vwUPCItemsWithCurrentPricing.Descri"
                        "ption, vwUPCItemsWithCurrentPricing.OnShelf, vwUPCItemsWithCurrentPricing.CurPri"
                        "ce AS Price, vwUPCItemsWithCurrentPricing.Cost, vwUPCItemsWithCurrentPricing.ID "
                        "AS ItemID FROM vwUPCItemsWithCurrentPricing ORDER BY vwUPCItemsWithCurrentPricin"
                        "g.LenID, vwUPCItemsWithCurrentPricing.UPCCode;"
                    ColumnWidths ="1440;2880;1440;1440;1440;1440"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =585
                            Height =210
                            Name ="Label32"
                            Caption ="Item ID"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =4
                    ListRows =25
                    ListWidth =10080
                    Left =1260
                    Top =180
                    Width =3720
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =1
                    TabIndex =9
                    BoundColumn =3
                    BackColor =16777215
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="ItemDescription"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByDesc.Description, lkupCluOrderItemByDesc.Cost, lkupCluO"
                        "rderItemByDesc.CurPrice, lkupCluOrderItemByDesc.ID AS ItemID FROM lkupCluOrderIt"
                        "emByDesc ORDER BY lkupCluOrderItemByDesc.Description;"
                    ColumnWidths ="2880;1440;1440;1440"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =540
                            Width =870
                            Height =210
                            Name ="Label33"
                            Caption ="Description"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =25
                    ListWidth =2880
                    Left =5040
                    Top =180
                    Width =720
                    Height =255
                    ColumnWidth =-2
                    ColumnOrder =2
                    TabIndex =10
                    BoundColumn =1
                    BackColor =16777215
                    ForeColor =0
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="Size"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupCluOrderItemByPkgSize.PkgSize, lkupCluOrderItemByPkgSize.ID AS ItemID"
                        " FROM lkupCluOrderItemByPkgSize; "
                    ColumnWidths ="1440;1440"
                    FontName ="Trebuchet MS"
                    ShortcutMenuBar ="KMDSPopup"

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4860
                            Top =540
                            Width =870
                            Height =210
                            Name ="Label34"
                            Caption ="Size"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9180
                    Top =960
                    ColumnWidth =-2
                    ColumnOrder =11
                    TabIndex =11
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8400
                            Top =960
                            Width =705
                            Height =240
                            Name ="Label21"
                            Caption ="ExtPrice:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9480
                    Top =1260
                    ColumnWidth =-2
                    ColumnOrder =13
                    TabIndex =12
                    Name ="ExtCost"
                    ControlSource ="ExtCost"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8580
                            Top =1260
                            Width =660
                            Height =240
                            Name ="Label22"
                            Caption ="ExtCost:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6120
                    Top =660
                    Width =660
                    ColumnWidth =0
                    ColumnOrder =16
                    TabIndex =13
                    Name ="Var"
                    ControlSource ="Var"

                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11580
                    Top =840
                    ColumnWidth =0
                    ColumnOrder =18
                    TabIndex =14
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10140
                            Top =600
                            Width =285
                            Height =240
                            Name ="Label37"
                            Caption ="ID:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =6840
                    Top =420
                    Width =1260
                    Height =255
                    ColumnWidth =630
                    ColumnOrder =4
                    TabIndex =15
                    BackColor =16777215
                    ForeColor =0
                    Name ="Front"
                    ControlSource ="CountFront"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =6840
                    Top =660
                    Width =1260
                    Height =255
                    ColumnWidth =630
                    ColumnOrder =5
                    TabIndex =16
                    BackColor =16777215
                    ForeColor =0
                    Name ="Back"
                    ControlSource ="CountBack"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =6840
                    Top =900
                    Width =1260
                    Height =255
                    ColumnWidth =825
                    ColumnOrder =6
                    TabIndex =17
                    BackColor =16777215
                    ForeColor =0
                    Name ="Display"
                    ControlSource ="CountDisplay"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =6840
                    Top =1140
                    Width =1260
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =7
                    TabIndex =18
                    BackColor =16777215
                    ForeColor =0
                    Name ="Overflow"
                    ControlSource ="CountOverflow"
                    StatusBarText ="Press {F8} to Search or Add New Record"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000640000000100000002000000ffffffff000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000002000000ffffffff010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =1320
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =900
                    Height =255
                    Name ="TotOnShelf"
                    ControlSource ="=Sum([OnShelf])"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Width =1260
                    Height =255
                    TabIndex =1
                    Name ="TotPhysicalCount"
                    ControlSource ="=Sum([PhysicalCount])"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2520
                    Width =720
                    Height =255
                    TabIndex =2
                    Name ="TotVarianceAmount"
                    ControlSource ="=Sum([VarianceAmount])"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    TabIndex =3
                    Name ="TotExtPrice"
                    ControlSource ="=Sum([ExtPrice])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3660
                            Top =60
                            Width =705
                            Height =240
                            Name ="Label28"
                            Caption ="ExtPrice:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9060
                    Top =120
                    TabIndex =4
                    Name ="TotExtCost"
                    ControlSource ="=Sum([ExtCost])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7620
                            Top =120
                            Width =660
                            Height =240
                            Name ="Label30"
                            Caption ="ExtCost:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =720
                    Width =1095
                    Height =405
                    TabIndex =5
                    Name ="cmdAddNew"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =720
                    Width =1185
                    Height =405
                    TabIndex =6
                    Name ="cmdSave"
                    Caption ="Save Record"
                    OnClick ="[Event Procedure]"

                End
            End
        End
    End
End
CodeBehindForm
' See "AuditSub.cls"
