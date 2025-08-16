Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13055
    DatasheetFontHeight =10
    ItemSuffix =148
    Left =390
    Top =420
    Right =13620
    Bottom =9420
    FrozenColumns =3
    HelpContextId =2200
    AfterInsert ="[Event Procedure]"
    Tag ="54 Club List"
    OrderBy ="[Clubs].[EMail] DESC, [Clubs].[City]"
    RecSrcDt = Begin
        0xee6d48d8e995e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT Club.* FROM Club ORDER BY Club.Name; "
    Caption ="Dispenser Setup and Change KMDS Software ©  KMDS Software ©  KMDS Software ©  KM"
        "DS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software "
        "©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
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
        Begin CustomControl
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
            Visible = NotDefault
            Height =240
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    ColumnOrder =0
                    Name ="Text62"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Width =615
                            Height =240
                            Name ="Label63"
                            Caption ="ID"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7740
            BackColor =-2147483633
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =1155
                    Top =7260
                    Width =1845
                    Height =300
                    TabIndex =7
                    BackColor =12632256
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =1155
                    LayoutCachedTop =7260
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =7560
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =600
                            Top =7260
                            Width =495
                            Height =240
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                            LayoutCachedLeft =600
                            LayoutCachedTop =7260
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =7500
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Width =12780
                    Height =7620
                    Name ="boxScreen"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =7620
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =10650
                    Top =7260
                    Width =1845
                    Height =300
                    TabIndex =6
                    BackColor =12632256
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="General Date"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =10650
                    LayoutCachedTop =7260
                    LayoutCachedWidth =12495
                    LayoutCachedHeight =7560
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =9960
                            Top =7260
                            Width =630
                            Height =240
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =9960
                            LayoutCachedTop =7260
                            LayoutCachedWidth =10590
                            LayoutCachedHeight =7500
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =4365
                    Top =180
                    Width =1050
                    Height =300
                    ColumnOrder =0
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4080
                            Top =180
                            Width =240
                            Height =285
                            Name ="ID_Label"
                            Caption ="ID"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ColumnCount =2
                    ListRows =30
                    ListWidth =1440
                    Left =1080
                    Top =180
                    Width =2865
                    Height =300
                    ColumnWidth =0
                    ColumnOrder =3
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.ID, Club.Name FROM Club ORDER BY Club.Name; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Width =930
                            Height =495
                            Name ="Label15"
                            Caption ="Select a Licensee"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =120
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =495
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    Left =6000
                    Top =180
                    Width =2865
                    Height =300
                    ColumnWidth =2685
                    ColumnOrder =1
                    TabIndex =2
                    Name ="txtName"
                    ControlSource ="Name"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5460
                            Top =180
                            Width =480
                            Height =285
                            Name ="Desc_Label"
                            Caption ="Name"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =7200
                    Width =12780
                    Height =420
                    Name ="Box70"
                    LayoutCachedLeft =120
                    LayoutCachedTop =7200
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =7620
                End
                Begin Tab
                    OverlapFlags =247
                    Left =120
                    Top =600
                    Width =12675
                    Height =6600
                    TabIndex =4
                    Name ="TabCtl98"

                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =12795
                    LayoutCachedHeight =7200
                    Begin
                        Begin Page
                            OverlapFlags =255
                            Left =255
                            Top =1005
                            Width =12405
                            Height =6060
                            Name ="Page99"
                            Caption ="Profile"
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
                            LayoutCachedLeft =255
                            LayoutCachedTop =1005
                            LayoutCachedWidth =12660
                            LayoutCachedHeight =7065
                            Begin
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =6255
                                    Width =10740
                                    Height =603
                                    Name ="ClubNotes"
                                    ControlSource ="ClubNotes"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =1260
                                    LayoutCachedTop =6255
                                    LayoutCachedWidth =12000
                                    LayoutCachedHeight =6858
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =495
                                            Top =6255
                                            Width =705
                                            Height =420
                                            Name ="Label45"
                                            Caption ="Club Notes:"
                                            LayoutCachedLeft =495
                                            LayoutCachedTop =6255
                                            LayoutCachedWidth =1200
                                            LayoutCachedHeight =6675
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2040
                                    Top =1080
                                    Width =2910
                                    TabIndex =1
                                    Name ="Licensee"
                                    ControlSource ="Licensee"
                                    OnEnter ="[Event Procedure]"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2040
                                    LayoutCachedTop =1080
                                    LayoutCachedWidth =4950
                                    LayoutCachedHeight =1320
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =510
                                            Top =1080
                                            Width =1470
                                            Height =255
                                            Name ="Label93"
                                            Caption =" OLCC Premise #:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =510
                                            LayoutCachedTop =1080
                                            LayoutCachedWidth =1980
                                            LayoutCachedHeight =1335
                                        End
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =11520
                                    Top =1620
                                    Width =1140
                                    Height =1080
                                    TabIndex =2
                                    Name ="cmdRefreshLicensee"
                                    Caption ="Refresh Licensee File"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =11520
                                    LayoutCachedTop =1620
                                    LayoutCachedWidth =12660
                                    LayoutCachedHeight =2700
                                    Overlaps =1
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =1440
                                    Width =2865
                                    TabIndex =3
                                    Name ="Contact"
                                    ControlSource ="Contact"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =4920
                                    LayoutCachedHeight =1680
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1185
                                            Top =1440
                                            Width =750
                                            Height =285
                                            Name ="Label22"
                                            Caption ="Contact:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1185
                                            LayoutCachedTop =1440
                                            LayoutCachedWidth =1935
                                            LayoutCachedHeight =1725
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =1711
                                    Width =2865
                                    TabIndex =4
                                    Name ="Address1"
                                    ControlSource ="Address1"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =1711
                                    LayoutCachedWidth =4920
                                    LayoutCachedHeight =1951
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1125
                                            Top =1711
                                            Width =810
                                            Height =285
                                            Name ="Label23"
                                            Caption ="Address1:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1125
                                            LayoutCachedTop =1711
                                            LayoutCachedWidth =1935
                                            LayoutCachedHeight =1996
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =1982
                                    Width =2865
                                    TabIndex =5
                                    Name ="Address2"
                                    ControlSource ="Address2"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =1982
                                    LayoutCachedWidth =4920
                                    LayoutCachedHeight =2222
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1125
                                            Top =1982
                                            Width =810
                                            Height =285
                                            Name ="Label24"
                                            Caption ="Address2:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1125
                                            LayoutCachedTop =1982
                                            LayoutCachedWidth =1935
                                            LayoutCachedHeight =2267
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =2253
                                    Width =2865
                                    TabIndex =6
                                    Name ="City"
                                    ControlSource ="City"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =2253
                                    LayoutCachedWidth =4920
                                    LayoutCachedHeight =2493
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1485
                                            Top =2253
                                            Width =450
                                            Height =285
                                            Name ="Label25"
                                            Caption ="City:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1485
                                            LayoutCachedTop =2253
                                            LayoutCachedWidth =1935
                                            LayoutCachedHeight =2538
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =3165
                                    Top =2503
                                    Width =645
                                    TabIndex =7
                                    Name ="Zip"
                                    ControlSource ="Zip"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =3165
                                    LayoutCachedTop =2503
                                    LayoutCachedWidth =3810
                                    LayoutCachedHeight =2743
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2775
                                            Top =2503
                                            Width =345
                                            Height =240
                                            Name ="Label27"
                                            Caption ="Zip:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =2775
                                            LayoutCachedTop =2503
                                            LayoutCachedWidth =3120
                                            LayoutCachedHeight =2743
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =2524
                                    Width =645
                                    TabIndex =8
                                    Name ="St"
                                    ControlSource ="St"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =2524
                                    LayoutCachedWidth =2700
                                    LayoutCachedHeight =2764
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1305
                                            Top =2524
                                            Width =270
                                            Height =240
                                            Name ="Label26"
                                            Caption ="St:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1305
                                            LayoutCachedTop =2524
                                            LayoutCachedWidth =1575
                                            LayoutCachedHeight =2764
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2055
                                    Top =2803
                                    TabIndex =9
                                    Name ="Phone1"
                                    ControlSource ="Phone1"
                                    FontName ="Trebuchet MS"
                                    InputMask ="!\\(999\") \"000\\-0000;;_"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =2803
                                    LayoutCachedWidth =3495
                                    LayoutCachedHeight =3043
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1260
                                            Top =2803
                                            Width =675
                                            Height =285
                                            Name ="Label28"
                                            Caption ="Phone1:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1260
                                            LayoutCachedTop =2803
                                            LayoutCachedWidth =1935
                                            LayoutCachedHeight =3088
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =4350
                                    Top =2820
                                    TabIndex =10
                                    Name ="Phone2"
                                    ControlSource ="Phone2"
                                    FontName ="Trebuchet MS"
                                    InputMask ="!\\(999\") \"000\\-0000;;_"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =4350
                                    LayoutCachedTop =2820
                                    LayoutCachedWidth =5790
                                    LayoutCachedHeight =3060
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =3810
                                            Top =2820
                                            Width =420
                                            Height =285
                                            Name ="Label29"
                                            Caption ="FAX:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =3810
                                            LayoutCachedTop =2820
                                            LayoutCachedWidth =4230
                                            LayoutCachedHeight =3105
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =6135
                                    Top =2820
                                    Width =245
                                    TabIndex =11
                                    Name ="ApplyDiscount"
                                    ControlSource ="ApplyDiscount"
                                    ControlTipText ="Does this club receive a discount?"

                                    LayoutCachedLeft =6135
                                    LayoutCachedTop =2820
                                    LayoutCachedWidth =6380
                                    LayoutCachedHeight =3060
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6360
                                            Top =2820
                                            Width =4680
                                            Height =240
                                            Name ="Label72"
                                            Caption ="Allow Discount MUST BE CHECKED TO RECEIVE DISCOUNT! "
                                            LayoutCachedLeft =6360
                                            LayoutCachedTop =2820
                                            LayoutCachedWidth =11040
                                            LayoutCachedHeight =3060
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =2
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2055
                                    Top =3240
                                    TabIndex =12
                                    ForeColor =16711680
                                    Name ="OuncesColumn"
                                    ControlSource ="OuncesColumn"
                                    Format ="Fixed"
                                    StatusBarText ="Order guides and invoices will give a Price Per Ounce"
                                    ControlTipText ="Order guides and invoices will give a Price Per Ounce"

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =3240
                                    LayoutCachedWidth =3495
                                    LayoutCachedHeight =3480
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3675
                                            Top =3240
                                            Width =5790
                                            Height =240
                                            Name ="Label127"
                                            Caption ="Price per x.xx ounce as shown in the Order Guide and Customer Invoice columns"
                                            LayoutCachedLeft =3675
                                            LayoutCachedTop =3240
                                            LayoutCachedWidth =9465
                                            LayoutCachedHeight =3480
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =4080
                                    Width =480
                                    TabIndex =13
                                    Name ="OrderGuides"
                                    ControlSource ="OrderGuides"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of ORDER GUIDES to print when saving an order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =4080
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =4320
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =885
                                            Top =4080
                                            Width =1080
                                            Height =255
                                            Name ="Label36"
                                            Caption ="OrderGuides:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =885
                                            LayoutCachedTop =4080
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =4335
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2640
                                    Top =4080
                                    Width =2115
                                    TabIndex =14
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="OrderGuideSort"
                                    ControlSource ="OrderGuideSort"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =2640
                                    LayoutCachedTop =4080
                                    LayoutCachedWidth =4755
                                    LayoutCachedHeight =4320
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =2
                                            Left =2970
                                            Top =3765
                                            Width =675
                                            Height =255
                                            Name ="Label90"
                                            Caption ="Sort By"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =2970
                                            LayoutCachedTop =3765
                                            LayoutCachedWidth =3645
                                            LayoutCachedHeight =4020
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5310
                                    Top =4157
                                    Width =300
                                    Height =180
                                    TabIndex =15
                                    Name ="OrderGuideByCategory"
                                    ControlSource ="OrderGuideByCategory"
                                    ControlTipText ="Order Guide will be grouped by SubCategories"

                                    LayoutCachedLeft =5310
                                    LayoutCachedTop =4157
                                    LayoutCachedWidth =5610
                                    LayoutCachedHeight =4337
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =4388
                                    Width =480
                                    TabIndex =16
                                    Name ="Invoices"
                                    ControlSource ="Invoices"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of INVOICES to print when saving an order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =4388
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =4628
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =735
                                            Top =4380
                                            Width =1230
                                            Height =255
                                            Name ="Label39"
                                            Caption ="OLCC Invoices:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =735
                                            LayoutCachedTop =4380
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =4635
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2640
                                    Top =4388
                                    Width =2115
                                    TabIndex =17
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="InvoiceListSort"
                                    ControlSource ="InvoiceListSort"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =2640
                                    LayoutCachedTop =4388
                                    LayoutCachedWidth =4755
                                    LayoutCachedHeight =4628
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =4690
                                    Width =480
                                    TabIndex =18
                                    Name ="StoreInvoices"
                                    ControlSource ="StoreInvoices"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of INVOICES to print when saving an order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =4690
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =4930
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =450
                                            Top =4680
                                            Width =1515
                                            Height =255
                                            Name ="Label123"
                                            Caption ="Store Invoices:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =450
                                            LayoutCachedTop =4680
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =4935
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2640
                                    Top =4690
                                    Width =2115
                                    TabIndex =19
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="StoreInvoiceSort"
                                    ControlSource ="StoreInvoiceSort"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =2640
                                    LayoutCachedTop =4690
                                    LayoutCachedWidth =4755
                                    LayoutCachedHeight =4930
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =4980
                                    Width =480
                                    TabIndex =20
                                    Name ="InvoicesCombined"
                                    ControlSource ="InvoicesCombined"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of COMBINED (Liquor and Related) INVOICES to print when saving an"
                                        " order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =4980
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =5220
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =435
                                            Top =4997
                                            Width =1530
                                            Height =255
                                            Name ="Label84"
                                            Caption ="CustomerInvoices:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =435
                                            LayoutCachedTop =4997
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =5252
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2640
                                    Top =4980
                                    Width =2115
                                    TabIndex =21
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="InvoiceCombinedListSort"
                                    ControlSource ="InvoiceCombinedListSort"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =2640
                                    LayoutCachedTop =4980
                                    LayoutCachedWidth =4755
                                    LayoutCachedHeight =5220
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5310
                                    Top =4997
                                    Width =300
                                    Height =180
                                    TabIndex =22
                                    Name ="InvoiceByCategory"
                                    ControlSource ="InvoiceByCategory"
                                    ControlTipText ="Invoices will be grouped by SubCategories"

                                    LayoutCachedLeft =5310
                                    LayoutCachedTop =4997
                                    LayoutCachedWidth =5610
                                    LayoutCachedHeight =5177
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =5280
                                    Width =480
                                    TabIndex =23
                                    Name ="Picklists"
                                    ControlSource ="Picklists"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of PICKLISTS to print when saving an order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =5280
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =5520
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1230
                                            Top =5280
                                            Width =735
                                            Height =255
                                            Name ="Label41"
                                            Caption ="Picklists:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1230
                                            LayoutCachedTop =5280
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =5535
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2640
                                    Top =5280
                                    Width =2115
                                    TabIndex =24
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="PickListSort"
                                    ControlSource ="PickListSort"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =2640
                                    LayoutCachedTop =5280
                                    LayoutCachedWidth =4755
                                    LayoutCachedHeight =5520
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =2085
                                    Top =5580
                                    Width =480
                                    TabIndex =25
                                    Name ="PosPicks"
                                    ControlSource ="PosPicks"
                                    ValidationRule ="Between 0 And 5"
                                    ValidationText ="This value must be Between 0 And 5"
                                    FontName ="Trebuchet MS"
                                    ControlTipText ="Default number of Point of Sale PICKLISTS print when saving an order."
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2085
                                    LayoutCachedTop =5580
                                    LayoutCachedWidth =2565
                                    LayoutCachedHeight =5820
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1155
                                            Top =5580
                                            Width =810
                                            Height =255
                                            Name ="Label43"
                                            Caption ="Pos Picks:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1155
                                            LayoutCachedTop =5580
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =5835
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    Left =2055
                                    Top =5880
                                    Width =2910
                                    TabIndex =26
                                    Name ="EMail"
                                    ControlSource ="EMail"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2055
                                    LayoutCachedTop =5880
                                    LayoutCachedWidth =4965
                                    LayoutCachedHeight =6120
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1395
                                            Top =5863
                                            Width =570
                                            Height =255
                                            Name ="Label92"
                                            Caption ="E-Mail:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1395
                                            LayoutCachedTop =5863
                                            LayoutCachedWidth =1965
                                            LayoutCachedHeight =6118
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6195
                                    Top =4035
                                    Width =3300
                                    TabIndex =27
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End
                                    Name ="Combo128"
                                    ControlSource ="ClubOrderGuideReport"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                                        "ReportForms.ReportType)=\"ClubOrderGuide\")) ORDER BY ReportForms.ReportName; "

                                    LayoutCachedLeft =6195
                                    LayoutCachedTop =4035
                                    LayoutCachedWidth =9495
                                    LayoutCachedHeight =4275
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =2
                                            Left =6165
                                            Top =3720
                                            Width =1395
                                            Height =255
                                            Name ="Label129"
                                            Caption ="Report format for"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =6165
                                            LayoutCachedTop =3720
                                            LayoutCachedWidth =7560
                                            LayoutCachedHeight =3975
                                        End
                                    End
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =5100
                                    Top =1440
                                    Width =6120
                                    Height =300
                                    FontSize =10
                                    ForeColor =1279872587
                                    Name ="lblMapQuest"
                                    Caption ="Click here for &Google Map"
                                    Tag ="http://maps.google.com/maps"
                                    ControlTipText ="http://maps.google.com/maps"
                                    HyperlinkAddress ="http://maps.google.com/maps"
                                    LayoutCachedLeft =5100
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =11220
                                    LayoutCachedHeight =1740
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =4875
                                    Top =3540
                                    Width =1080
                                    Height =480
                                    Name ="Label75"
                                    Caption ="Group By SubCategory"
                                    FontName ="Trebuchet MS"
                                    LayoutCachedLeft =4875
                                    LayoutCachedTop =3540
                                    LayoutCachedWidth =5955
                                    LayoutCachedHeight =4020
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =1935
                                    Top =3780
                                    Width =645
                                    Height =240
                                    Name ="Label38"
                                    Caption ="Copies"
                                    FontName ="Trebuchet MS"
                                    LayoutCachedLeft =1935
                                    LayoutCachedTop =3780
                                    LayoutCachedWidth =2580
                                    LayoutCachedHeight =4020
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =675
                                    Top =3240
                                    Width =1245
                                    Height =240
                                    Name ="Label126"
                                    Caption ="Ounces Column:"
                                    LayoutCachedLeft =675
                                    LayoutCachedTop =3240
                                    LayoutCachedWidth =1920
                                    LayoutCachedHeight =3480
                                End
                                Begin OptionButton
                                    OverlapFlags =255
                                    Left =6960
                                    Top =1860
                                    TabIndex =28
                                    Name ="optObsolete"
                                    ControlSource ="NonLiquor"

                                    LayoutCachedLeft =6960
                                    LayoutCachedTop =1860
                                    LayoutCachedWidth =7220
                                    LayoutCachedHeight =2100
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =1
                                            Left =7260
                                            Top =1920
                                            Width =2640
                                            Height =585
                                            ForeColor =255
                                            Name ="lblObsolete"
                                            Caption ="This Licensee is OBSOLETE and will not appear in Create Order List"
                                            LayoutCachedLeft =7260
                                            LayoutCachedTop =1920
                                            LayoutCachedWidth =9900
                                            LayoutCachedHeight =2505
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =6840
                                    Top =1800
                                    Width =3360
                                    Height =960
                                    Name ="Box135"
                                    LayoutCachedLeft =6840
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =10200
                                    LayoutCachedHeight =2760
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =9120
                                    Top =5865
                                    Height =255
                                    TabIndex =29
                                    Name ="FintechAccount"
                                    ControlSource ="FintechAccount"

                                    LayoutCachedLeft =9120
                                    LayoutCachedTop =5865
                                    LayoutCachedWidth =10560
                                    LayoutCachedHeight =6120
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =7080
                                            Top =5685
                                            Width =1995
                                            Height =435
                                            Name ="Label136"
                                            Caption ="Fintech Account            \"0\"   = NOT Fintech"
                                            LayoutCachedLeft =7080
                                            LayoutCachedTop =5685
                                            LayoutCachedWidth =9075
                                            LayoutCachedHeight =6120
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =255
                            Top =1005
                            Width =12405
                            Height =6060
                            Name ="Page100"
                            Caption ="Order Guide"
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
                            LayoutCachedLeft =255
                            LayoutCachedTop =1005
                            LayoutCachedWidth =12660
                            LayoutCachedHeight =7065
                            Begin
                                Begin Subform
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =855
                                    Top =1020
                                    Width =8640
                                    Height =5820
                                    Name ="ClubOrderGuides"
                                    SourceObject ="Form.ClubOrderGuides"
                                    LinkChildFields ="ClubID"
                                    LinkMasterFields ="ID"

                                    LayoutCachedLeft =855
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =9495
                                    LayoutCachedHeight =6840
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    Left =10470
                                    Top =2085
                                    Width =606
                                    Height =426
                                    TabIndex =1
                                    Name ="cmdExportToExcel"
                                    Caption ="Command88"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Report to File"

                                    LayoutCachedLeft =10470
                                    LayoutCachedTop =2085
                                    LayoutCachedWidth =11076
                                    LayoutCachedHeight =2511
                                    Overlaps =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =11130
                                            Top =2085
                                            Width =1320
                                            Height =405
                                            Name ="Label89"
                                            Caption ="Export Order Guide to Excel "
                                            LayoutCachedLeft =11130
                                            LayoutCachedTop =2085
                                            LayoutCachedWidth =12450
                                            LayoutCachedHeight =2490
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =10650
                                    Top =2805
                                    Width =245
                                    TabIndex =2
                                    Name ="MaintainParLevels"
                                    ControlSource ="MaintainParLevels"

                                    LayoutCachedLeft =10650
                                    LayoutCachedTop =2805
                                    LayoutCachedWidth =10895
                                    LayoutCachedHeight =3045
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =10880
                                            Top =2805
                                            Width =1470
                                            Height =240
                                            Name ="Label74"
                                            Caption ="Maintain Par Levels"
                                            LayoutCachedLeft =10880
                                            LayoutCachedTop =2805
                                            LayoutCachedWidth =12350
                                            LayoutCachedHeight =3045
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =10395
                                    Top =3420
                                    Width =900
                                    Height =720
                                    TabIndex =3
                                    ForeColor =-2147483630
                                    Name ="cmdPalmFrom"
                                    Caption ="Command56"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                                        0x0707070707070707070707070707070707000700000000000000000000000000 ,
                                        0x0000070707070707070707070707070701f800ffffffffffffffffffffffffff ,
                                        0xffff000707070707070707000000000701f800ffffffffffffffffffffffffff ,
                                        0xffff000707070707070700f901f9010001f800f8f8f8f8f8f8f8f8f8fff800f8 ,
                                        0xffff0007070707070700f901f901000001f8000000000000ffffffff0000ffff ,
                                        0xffff00070707070700f901f9f900ffff0000ffffffffffff0000000001000707 ,
                                        0x070700070707070700f9f9f900ffffffffffffffffffffff0707070000ffffff ,
                                        0xffff00070707070700f9f9f900ffffffffffffffffff00000000000007070707 ,
                                        0x070700070707070700f9f9f900ffffffffffffffff0700ff00fb0000ffffffff ,
                                        0xfff80007070707070700f9f900ffffffffffffffff070000fb00070007070707 ,
                                        0x0707000707070707070700000000ffffffffffffff0700fb0007ff00fffffff8 ,
                                        0xf8f80007070707070707070707070000ffffffffff00fb0007ff000707070707 ,
                                        0x0707000707070707070707070707070700ffffff00fb0007ff00fffffff8f8f8 ,
                                        0xf8f8000707070707070707070707070701000000fb00ffff0007070707070707 ,
                                        0x0707000707070707070707070707070701f800fb00000000fffffff8f8f8f8f8 ,
                                        0xf8f800070707070707070707070707070100fb00070707070707070707070707 ,
                                        0x0707000707070707070707070707070700fb00fffffffffffff8f8f8f8f8f8f8 ,
                                        0xf8f80007070707070707070707070700fb000007070707070707070707070707 ,
                                        0x070700070707070707070707070700fb00f800ffffff00000000000000000000 ,
                                        0x0000000000000000070707070700fb0001f80007070700f9f9f9f9f9f9f9f9f9 ,
                                        0xf9f9f9f9f9f9f900070707070700000701f800ffff0005f905f905f905f905f9 ,
                                        0x05f905f905f90007070707070707070701f800070700f905f905f905f905f905 ,
                                        0xf905f905f9050007070707070707070701f800ff00f905f905f905f905f905f9 ,
                                        0x05f905f905000707070707070707070701f800070005f905f905f905f905f905 ,
                                        0xf905f905f9000707070707070707070701f8000005f905f905f905f905f905f9 ,
                                        0x05f905f900070707070707070707070701f80000f905f905f905f905f905f905 ,
                                        0xf905f905000707070707070707070707010700f905f905f905f905f905f905f9 ,
                                        0x05f9050007070707070707070707070707010005000001000001000001000001 ,
                                        0x0000010007070707070707070707070707010100f80000f80000f80000f80000 ,
                                        0xf800000707070707070707070707070707070707000007000007000007000007 ,
                                        0x0000070707070707070707070707070707070707070707070707070707070707 ,
                                        0x0707070707070707
                                    End
                                    FontName ="Tahoma"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Run Query"

                                    LayoutCachedLeft =10395
                                    LayoutCachedTop =3420
                                    LayoutCachedWidth =11295
                                    LayoutCachedHeight =4140
                                    Overlaps =1
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =11355
                                            Top =3420
                                            Width =1020
                                            Height =660
                                            BackColor =16777215
                                            ForeColor =0
                                            Name ="Label57"
                                            Caption ="Receive from Palm (Tab)"
                                            FontName ="Tahoma"
                                            LayoutCachedLeft =11355
                                            LayoutCachedTop =3420
                                            LayoutCachedWidth =12375
                                            LayoutCachedHeight =4080
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =255
                            Top =1005
                            Width =12405
                            Height =6060
                            Name ="Page101"
                            Caption ="Previously Sold Purchase Orders"
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
                            LayoutCachedLeft =255
                            LayoutCachedTop =1005
                            LayoutCachedWidth =12660
                            LayoutCachedHeight =7065
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1935
                                    Top =1200
                                    Width =10140
                                    Height =5670
                                    Name ="ClubsPreviousOrderSub"
                                    SourceObject ="Form.ClubsPreviousOrderSub"
                                    LinkChildFields ="clubid"
                                    LinkMasterFields ="id"

                                    LayoutCachedLeft =1935
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =12075
                                    LayoutCachedHeight =6870
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =315
                                    Top =1200
                                    Width =1440
                                    Height =960
                                    Name ="Label125"
                                    Caption ="Click on Invoice Number to View Invoice"
                                    LayoutCachedLeft =315
                                    LayoutCachedTop =1200
                                    LayoutCachedWidth =1755
                                    LayoutCachedHeight =2160
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =255
                                    Top =2340
                                    Width =1620
                                    Height =720
                                    TabIndex =1
                                    Name ="cmdReprintInvoice"
                                    Caption ="Reprint InvoiceId #13035"
                                    OnClick ="[Event Procedure]"
                                    Tag ="421853"
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =255
                                    LayoutCachedTop =2340
                                    LayoutCachedWidth =1875
                                    LayoutCachedHeight =3060
                                    Overlaps =1
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5880
                    Top =7260
                    Width =2535
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="txtTotalSales"
                    ControlSource ="=ClubsPreviousOrderSub.Form!SumTotalPrice"

                    LayoutCachedLeft =5880
                    LayoutCachedTop =7260
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =7560
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4440
                            Top =7260
                            Width =1215
                            Height =255
                            Name ="Label104"
                            Caption ="Total Sales"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =7260
                            LayoutCachedWidth =5655
                            LayoutCachedHeight =7515
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =25
                    ListWidth =10080
                    Left =9780
                    Top =180
                    Width =3120
                    Height =300
                    ColumnWidth =2055
                    TabIndex =3
                    BoundColumn =1
                    Name ="cmbLicensee"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT  Licensee.TradeName, Licensee.PremisesNumber, Licensee.Street, Licensee.C"
                        "ity, Licensee.Zip FROM Licensee ORDER BY Licensee.TradeName;"
                    ColumnWidths ="4515;1440;2340;1124;720"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =8940
                            Top =180
                            Width =765
                            Height =285
                            Name ="Licensee_Label"
                            Caption ="Licensee"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =9555
                    Top =4560
                    Width =690
                    TabIndex =8
                    BackColor =12632256
                    Name ="OrderCount"
                    ControlSource ="OrderCount"
                    DefaultValue ="1"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =9555
                    LayoutCachedTop =4560
                    LayoutCachedWidth =10245
                    LayoutCachedHeight =4800
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =8460
                            Top =4560
                            Width =1020
                            Height =285
                            Name ="Label31"
                            Caption ="OrderCount:"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8460
                            LayoutCachedTop =4560
                            LayoutCachedWidth =9480
                            LayoutCachedHeight =4845
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =8820
                    Top =660
                    Width =3845
                    Height =358
                    TabIndex =9
                    Name ="optLicenseeSort"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    LayoutCachedLeft =8820
                    LayoutCachedTop =660
                    LayoutCachedWidth =12665
                    LayoutCachedHeight =1018
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =8940
                            Top =540
                            Width =840
                            Height =240
                            Name ="Label142"
                            Caption ="Lookup By"
                            LayoutCachedLeft =8940
                            LayoutCachedTop =540
                            LayoutCachedWidth =9780
                            LayoutCachedHeight =780
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =10020
                            Top =748
                            OptionValue =1
                            Name ="Option144"
                            LayoutCachedLeft =10020
                            LayoutCachedTop =748
                            LayoutCachedWidth =10280
                            LayoutCachedHeight =988
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =10250
                                    Top =720
                                    Width =510
                                    Height =240
                                    Name ="Label145"
                                    Caption ="Name"
                                    LayoutCachedLeft =10250
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =10760
                                    LayoutCachedHeight =960
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =10860
                            Top =748
                            OptionValue =2
                            Name ="Option146"
                            LayoutCachedLeft =10860
                            LayoutCachedTop =748
                            LayoutCachedWidth =11120
                            LayoutCachedHeight =988
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =11090
                                    Top =720
                                    Width =645
                                    Height =240
                                    Name ="Label147"
                                    Caption ="Number"
                                    LayoutCachedLeft =11090
                                    LayoutCachedTop =720
                                    LayoutCachedWidth =11735
                                    LayoutCachedHeight =960
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =8280
                    Top =120
                    Width =666
                    Height =606
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="macKMDSDelete"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8280
                    LayoutCachedTop =120
                    LayoutCachedWidth =8946
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =8318
                            Top =720
                            Width =555
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8318
                            LayoutCachedTop =720
                            LayoutCachedWidth =8873
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =7200
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =1
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
                    OnClick ="macKMDSSave"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8070700000000000000000000000000000000000000000000 ,
                        0x0000000000f8f8f8070007070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700000000000000000000000000000000 ,
                        0x000007070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070000000000000000000000000000000000000000 ,
                        0x000007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8070007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff00070707000707070000000000000000000000000000000000000000000000 ,
                        0x0000000000070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7200
                    LayoutCachedTop =120
                    LayoutCachedWidth =7866
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =7267
                            Top =720
                            Width =420
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7267
                            LayoutCachedTop =720
                            LayoutCachedWidth =7687
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =6120
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =2
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
                    OnClick ="macKMDSUndo"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070700 ,
                        0xfb01fb01000707070707070707070707070707070707070707070707070700fb ,
                        0x01fb01000107070707070707070707070707070707070707070707070700fb01 ,
                        0xfb010001fb070707070707070707070707070707070707070707070700fb01fb ,
                        0x010001fb010707070707070707070707070707070707070707070700fb01fb01 ,
                        0x0001fb01fb07070707070707070707070707070707070707070700fb01fb0100 ,
                        0x01fb01fb00070707070707070707070707070707070707070700fb01fb010001 ,
                        0xfb01fb00fb0707070707070707070707070707070707070700fb01fb010001fb ,
                        0x01fb00fb0007070707070707070707070707070707070700000000010001fb01 ,
                        0xfb00fb0007070707070707070707070707070707070700020202000001fb01fb ,
                        0x00fb00070707070707070707070707070707070707000000000202020001fb00 ,
                        0xfb0007070707070707070707070707070707070700fb01fb00000202000000fb ,
                        0x0007070707070707070707070707070707070700fb01fb01fb0100020200fb00 ,
                        0x07070707070707070707070707070707070700fb01fb01fb01fb000002000007 ,
                        0x07070707070707070707070707070707070000000001fb01fb01fb0002000707 ,
                        0x0707070707070707070707070707070700020202000001fb01fb010000070707 ,
                        0x070707070707070707070707070707000000000202020001fb01fb0007070707 ,
                        0x07070707070707070707070707070007fd0700000202000001fb000707070707 ,
                        0x070707070707070707070707070007fd07fd07fd00020200fb00070707070707 ,
                        0x070707070707070707070707000000000007fd07000002000007070707070707 ,
                        0x07070707070707070707070700fdfffdff0007fd070002000707070707070707 ,
                        0x070707070707070707070700fdfffdfffd00fd07fd0000070707070707070707 ,
                        0x070707070707070707070700fffdfffdff0000fd070007070707070707070707 ,
                        0x070707070707070707070700fdfffdfffdff0007000707070707070707070707 ,
                        0x070707070707070707070700fffdfffdfffd0000070707070707070707070707 ,
                        0x07070707070707070707070700fffdfffdff0007070707070707070707070707 ,
                        0x07070707070707070707070700fdfffdff000707070707070707070707070707 ,
                        0x0707070707070707070707070700000000070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="Trebuchet MS"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =110
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =6786
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =6187
                            Top =720
                            Width =450
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6187
                            LayoutCachedTop =720
                            LayoutCachedWidth =6637
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =11520
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =3
                    Name ="Command0"
                    Caption ="Command64"
                    OnClick ="macKMDSClose"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =11520
                    LayoutCachedTop =120
                    LayoutCachedWidth =12186
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =11520
                            Top =720
                            Width =885
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =11520
                            LayoutCachedTop =720
                            LayoutCachedWidth =12405
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =5040
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =4
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5040
                    LayoutCachedTop =120
                    LayoutCachedWidth =5706
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =5048
                            Top =720
                            Width =705
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5048
                            LayoutCachedTop =720
                            LayoutCachedWidth =5753
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =9360
                    Top =120
                    Width =666
                    Height =606
                    FontSize =9
                    TabIndex =5
                    ForeColor =-2147483630
                    Name ="cmdPrintPreview"
                    Caption ="Show All Agents"
                    OnClick ="macKMDSPrintPreview"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000030507050203040ff ,
                        0x203850500000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff504840ff305060ff4078c0ff ,
                        0x304860ff3050605000000000b0a090ffffffffffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ff908880ff305070ff3080d0ff4098e0ff ,
                        0x50b0f0ff506870ff00000000b0a090fffffffffffffffffffff8fffff0f0f0ff ,
                        0xd0d8d0ff909890ff706860ff606860ff505860ff407090ff40a0e0ff60c8ffff ,
                        0x7090a0ff7088905000000000b0a090ffffffffffffffffffffffffffe0e0e0ff ,
                        0x909090ffb0a8a0ffd0c0b0ffd0b0a0ff807870ff505850ff6090b0ff7098b0ff ,
                        0x708890800000000000000000b0a090ffffffffffffffffffffffffffb0b0b0ff ,
                        0xc0b8b0fffff0e0ffffe8e0fff0d8c0fff0d0b0ff807870ff6058609070889080 ,
                        0x000000000000000000000000c0a890ffffffffffffffffffffffffffa09890ff ,
                        0xf0e8e0fffff8f0fffff0f0ffffe8e0fff0d8d0ffd0b0a0ff505850e000000000 ,
                        0x000000000000000000000000c0a8a0ffffffffffffffffffffffffffa0a090ff ,
                        0xf0e8e0fffffffffffff8f0fffff0f0ffffe8e0ffe0c0b0ff606060e000000000 ,
                        0x000000000000000000000000c0b0a0ffffffffffffffffffffffffffc0c8c0ff ,
                        0xc0c0c0fffffffffffffffffffff8f0fffff0e0ffb0a090ff6060608000000000 ,
                        0x000000000000000000000000d0b0a0fffffffffffffffffffffffffff0f8ffff ,
                        0xc0b8b0ffc0c0c0fff0e8e0fff0e8e0ffb0b0a0ff707060ff8070601000000000 ,
                        0x000000000000000000000000d0b8a0ffffffffffffffffffffffffffffffffff ,
                        0xf0f8ffffc0c8c0ffa0a090ff909080ff909090ff605040ff0000000000000000 ,
                        0x000000000000000000000000d0b8b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffb0a090ff604830ff604830ff604830ff0000000000000000 ,
                        0x000000000000000000000000d0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a890ffd0c8c0ff604830ffd0b0a0900000000000000000 ,
                        0x000000000000000000000000e0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a8a0ff604830ffd0b0a090000000000000000000000000 ,
                        0x000000000000000000000000e0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b0a0ffd0b0a09000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9360
                    LayoutCachedTop =120
                    LayoutCachedWidth =10026
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =9360
                            Top =720
                            Width =1095
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label61"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =9360
                            LayoutCachedTop =720
                            LayoutCachedWidth =10455
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =10440
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =6
                    Name ="cmdDataSheet"
                    Caption ="Command59"
                    OnClick ="macKMDSDatasheet"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000806860ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x0000000000000000907860ffd0c8c0ff805840fff0f0f0ffe0d8d0ffd0d0d0ff ,
                        0xd0c8c0ffd0c0c0ffc0b8b0ffc0b0b0ffb0a8a0ffb0a8a0ffa09890ff604830ff ,
                        0x0000000000000000a08070ffd0c8d0ff805840fff0f0f0ffb09080ffa08070ff ,
                        0xe0e0d0ff907050ff906850ffd0c8c0ff906850ff906850ffb09890ff604830ff ,
                        0x0000000000000000a09080ffd0d0d0ff805840fffff8f0fff0f0f0fff0f0f0ff ,
                        0xf0e8e0ffe0e0e0ffe0d8d0ffe0d0d0ffd0c8c0ffd0c0b0ffb0a8a0ff604830ff ,
                        0x0000000000000000b0a090ffd0d0d0ff906040fffff8ffffc0a8a0ffc09890ff ,
                        0xf0f0f0ffa08070ffa07860ffe0d8d0ff906850ff906850ffc0b0b0ff604830ff ,
                        0x0000000000000000c0a8a0ffe0d8e0ff906850fffffffffffff8fffffff8f0ff ,
                        0xf0f0f0fff0f0f0fff0e8e0ffe0e0e0ffe0d8d0ffe0d8d0ffc0b8b0ff604830ff ,
                        0x0000000000000000c0b0a0ffe0e0e0ffa07050ffffffffffc0a8a0ffc0a090ff ,
                        0xfff8f0ffc09890ffb09080fff0e8e0ffa08060ffa07860ffd0c8c0ff604830ff ,
                        0x0000000000000000c0b8a0ffe0e0e0ffa07860ffffffffffffffffffffffffff ,
                        0xfff8fffffff8f0fff0f0f0fff0f0f0fff0e8e0ffe0e0e0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0b8b0fff0e8f0ffb08060ffffffffffc0a890ffc0a090ff ,
                        0xffffffffc0a090ffc0a090fff0f0f0ffb09080ffb09080ffe0d8d0ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffb08870ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffff8fffffff8f0fff0f0f0fff0f0f0ffe0e0e0ff604830ff ,
                        0x0000000000000000d0c0b0ffc09080ffb09070ffb08870ffa08060ffa07860ff ,
                        0x907050ff906850ff906040ff806040ff805840ff805840ff805840ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffc09080fff0f0f0fff0e8f0ffe0e8e0ff ,
                        0xe0e0e0ffe0e0e0ffe0d8e0ffd0d8d0ffd0d0d0ffd0d0d0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0b8b0ffc0b0a0ff ,
                        0xc0b0a0ffc0a8a0ffb0a090ffb09890ffa09080ffa08870ff908070ff907860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10440
                    LayoutCachedTop =120
                    LayoutCachedWidth =11106
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =10440
                            Top =720
                            Width =915
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10440
                            LayoutCachedTop =720
                            LayoutCachedWidth =11355
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =120
                    Top =60
                    Width =12780
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =1020
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =240
                    Top =120
                    Width =606
                    Height =426
                    TabIndex =7
                    Name ="cmdExportAllToExcel"
                    Caption ="Command88"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Report to File"

                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =846
                    LayoutCachedHeight =546
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =900
                            Top =120
                            Width =1500
                            Height =840
                            Name ="Label95"
                            Caption ="Export ALL Order Guide to OO Excel.Exported to K:\\KMDS\\Clubs\\  "
                            LayoutCachedLeft =900
                            LayoutCachedTop =120
                            LayoutCachedWidth =2400
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    AccessKey =71
                    Left =2820
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =8
                    Name ="cmdMailOrderGuide"
                    Caption ="Command109"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x070707070707070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000ff000000fffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffb000000000000000000000000fffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000000000000000fffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707c2686d136a13000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Mail Report"
                    UnicodeAccessKey =71

                    LayoutCachedLeft =2820
                    LayoutCachedTop =120
                    LayoutCachedWidth =3396
                    LayoutCachedHeight =696
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =3390
                            Top =120
                            Width =930
                            Height =840
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label110"
                            Caption ="eMail Excel Order &Guide"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3390
                            LayoutCachedTop =120
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =960
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Clubs.cls"
