Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10860
    DatasheetFontHeight =10
    ItemSuffix =172
    Top =600
    Right =12645
    Bottom =10710
    HelpContextId =7400
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xcfcef26bcf9de240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="CashSys"
    Caption =" KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =6540
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =2175
                    Top =240
                    Width =420
                    Height =255
                    ColumnWidth =1440
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedLeft =2175
                    LayoutCachedTop =240
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =495
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =120
                            Top =240
                            Width =1995
                            Height =240
                            Name ="ID_Label"
                            Caption ="Setting up a New Record"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =2175
                    Top =900
                    Width =1935
                    Height =255
                    ColumnWidth =1935
                    TabIndex =4
                    Name ="Desc"
                    ControlSource ="Description"
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
                            OverlapFlags =93
                            TextAlign =3
                            Left =510
                            Top =900
                            Width =1605
                            Height =240
                            Name ="Desc_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =10800
                    Height =6480
                    Name ="boxScreen"
                End
                Begin ComboBox
                    OverlapFlags =247
                    ColumnCount =2
                    ListRows =30
                    ListWidth =4680
                    Left =2175
                    Top =240
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName, CashSys.Description FROM CashSys ORDER "
                        "BY CashSys.ID; "
                    ColumnWidths ="720;2160;2160"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =2175
                    LayoutCachedTop =240
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =240
                            Top =240
                            Width =1875
                            Height =240
                            Name ="Label15"
                            Caption ="Select a Computer"
                            LayoutCachedLeft =240
                            LayoutCachedTop =240
                            LayoutCachedWidth =2115
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =255
                    Left =105
                    Top =1260
                    Width =10710
                    Height =4860
                    TabIndex =5
                    Name ="tabCashsys"

                    Begin
                        Begin Page
                            OverlapFlags =247
                            Left =240
                            Top =1665
                            Width =10440
                            Height =4320
                            Name ="Page48"
                            Caption ="POS Periphials"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3435
                                    Top =1995
                                    Name ="DisplayPort"
                                    ControlSource ="DisplayPort"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2415
                                            Top =1995
                                            Width =960
                                            Height =240
                                            Name ="Label52"
                                            Caption ="Display Port:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    ListWidth =1440
                                    Left =6075
                                    Top =1995
                                    TabIndex =1
                                    Name ="ReportsDir"
                                    ControlSource ="ReportsDir"
                                    RowSourceType ="Value List"
                                    RowSource ="\"Logic Controls\";\"IEE\";\"IEEPT\""
                                    ColumnWidths ="1440"
                                    DefaultValue ="\"Logic Controls\""

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4935
                                            Top =1995
                                            Width =1065
                                            Height =245
                                            Name ="Manufacturer_Label"
                                            Caption ="Manufacturer"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3420
                                    Top =2295
                                    Width =4575
                                    TabIndex =2
                                    Name ="DisplayDefault"
                                    ControlSource ="DisplayDefault"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2175
                                            Top =2295
                                            Width =1185
                                            Height =240
                                            Name ="Label53"
                                            Caption ="Display Default:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3435
                                    Top =2940
                                    TabIndex =3
                                    Name ="PrinterPort"
                                    ControlSource ="PrinterPort"
                                    OnExit ="[Event Procedure]"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2520
                                            Top =2940
                                            Width =855
                                            Height =240
                                            Name ="Label55"
                                            Caption ="PrinterPort:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =6105
                                    Top =2955
                                    TabIndex =4
                                    Name ="DrawerPort"
                                    ControlSource ="DrawerPort"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =5115
                                            Top =2955
                                            Width =930
                                            Height =240
                                            Name ="Label54"
                                            Caption ="Drawer Port"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3465
                                    Top =3255
                                    TabIndex =5
                                    Name ="chkEpsonPrinterInstalled"
                                    ControlSource ="EpsonPrinterInstalled"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1755
                                            Top =3255
                                            Width =1650
                                            Height =240
                                            Name ="Label59"
                                            Caption ="Epson Printer Installed"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3465
                                    Top =3555
                                    TabIndex =6
                                    Name ="chkEpsonPrintertSlit"
                                    ControlSource ="EpsonPrintertSlit"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2010
                                            Top =3555
                                            Width =1395
                                            Height =240
                                            Name ="Label61"
                                            Caption ="Cut Receipt"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3435
                                    Top =3855
                                    TabIndex =11
                                    Name ="chkStoreSmode"
                                    ControlSource ="StoreSmode"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =1815
                                            Top =3855
                                            Width =1545
                                            Height =240
                                            Name ="Label96"
                                            Caption ="Print Narrow Receipt"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =3435
                                    Top =4155
                                    Width =4560
                                    TabIndex =7
                                    Name ="InvoiceTrailer"
                                    ControlSource ="InvoiceTrailer"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =2295
                                            Top =4155
                                            Width =1080
                                            Height =240
                                            Name ="Label70"
                                            Caption ="InvoiceTrailer:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =795
                                    Top =1875
                                    Width =7200
                                    Height =780
                                    Name ="boxCustDisplay"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =795
                                    Top =2835
                                    Width =7200
                                    Height =1740
                                    Name ="Box77"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =795
                                    Top =4695
                                    Width =7200
                                    Height =840
                                    Name ="Box78"
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3420
                                    Top =4875
                                    Width =735
                                    Height =255
                                    TabIndex =8
                                    Name ="ScannerComPort"
                                    ControlSource ="ScannerComPort"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1995
                                            Top =4875
                                            Width =1365
                                            Height =255
                                            Name ="Label74"
                                            Caption ="Com Port Number:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3435
                                    Top =5115
                                    Height =255
                                    TabIndex =9
                                    Name ="ScannerSettings"
                                    ControlSource ="ScannerSettings"
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
                                            Left =5010
                                            Top =5115
                                            Width =1230
                                            Height =240
                                            Name ="Label100"
                                            Caption ="e.g., 9600,N,8,1"
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =915
                                    Top =1815
                                    Width =1290
                                    Height =240
                                    Name ="Label97"
                                    Caption ="Customer Display"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =915
                                    Top =2775
                                    Width =1125
                                    Height =240
                                    Name ="Label98"
                                    Caption ="Receipt Printer"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =915
                                    Top =4635
                                    Width =1320
                                    Height =240
                                    Name ="Label99"
                                    Caption ="Barcode Scanner"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =2025
                                    Top =5115
                                    Width =1335
                                    Height =255
                                    Name ="Label75"
                                    Caption ="Com port Mode :"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =240
                            Top =1665
                            Width =10440
                            Height =4320
                            Name ="Page49"
                            Caption ="Settings and Flags"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3060
                                    Top =1890
                                    TabIndex =1
                                    Name ="chkPrintReceipt"
                                    ControlSource ="PrintReceipt"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =990
                                            Top =1890
                                            Width =2010
                                            Height =240
                                            Name ="Label63"
                                            Caption ="Print Receipt for every Sale"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7800
                                    Top =1890
                                    Name ="Check111"
                                    ControlSource ="ShowAvail"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =4680
                                            Top =1890
                                            Width =3060
                                            Height =630
                                            Name ="Label112"
                                            Caption ="Display on-shelf during sale.. (a network query is peformed for each line to dis"
                                                "play current onhands)."
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3060
                                    Top =2250
                                    TabIndex =2
                                    Name ="chkAllowExactCh"
                                    ControlSource ="AllowExactCh"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1350
                                            Top =2250
                                            Width =1650
                                            Height =240
                                            Name ="Label65"
                                            Caption ="Allow Exact Change"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3060
                                    Top =2610
                                    TabIndex =3
                                    Name ="chkAllowEscape"
                                    ControlSource ="AllowEscape"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =840
                                            Top =2610
                                            Width =2160
                                            Height =240
                                            Name ="Label67"
                                            Caption ="Uses aTouch Screen Monitor"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =7740
                                    Top =2850
                                    Name ="Check113"
                                    ControlSource ="PlaySounds"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =4680
                                            Top =2850
                                            Width =3030
                                            Height =630
                                            Name ="Label114"
                                            Caption ="Provide feedback with sound as each line is recorded and a warning sound when a "
                                                "'not on file' is entered."
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3060
                                    Top =2970
                                    TabIndex =4
                                    Name ="chkReprintnegsales"
                                    ControlSource ="Reprintnegsales"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1230
                                            Top =2970
                                            Width =1770
                                            Height =240
                                            Name ="Label69"
                                            Caption ="Reprint Negative Sales"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =3060
                                    Top =3330
                                    Name ="Check101"
                                    ControlSource ="PrintZReports"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1230
                                            Top =3330
                                            Width =1770
                                            Height =240
                                            Name ="Label102"
                                            Caption ="Print Z Reports"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =2790
                                    Top =3720
                                    TabIndex =5
                                    Name ="ReceiptSpaces"
                                    ControlSource ="ReceiptSpaces"
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
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =1500
                                            Top =3720
                                            Width =1260
                                            Height =240
                                            Name ="Label73"
                                            Caption ="Receipt Spaces:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =660
                                    Top =1770
                                    Width =3840
                                    Height =1860
                                    Name ="Box80"
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =660
                                    Top =3630
                                    Width =3840
                                    Height =1905
                                    Name ="Box81"
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =6870
                                    Top =3720
                                    Name ="Check117"
                                    ControlSource ="ShowRemoveItem"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =4605
                                            Top =3705
                                            Width =2205
                                            Height =240
                                            Name ="Label118"
                                            Caption ="Display/Enable Remove Item"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2790
                                    Top =4080
                                    TabIndex =6
                                    Name ="LastInvoiceNo"
                                    ControlSource ="LastInvoiceNo"
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
                                            Left =1515
                                            Top =4080
                                            Width =1245
                                            Height =240
                                            Name ="Label72"
                                            Caption ="Last Invoice No:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =6870
                                    Top =4080
                                    ColumnWidth =3030
                                    Name ="Check119"
                                    ControlSource ="ShowDupeItem"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =4590
                                            Top =4065
                                            Width =2220
                                            Height =240
                                            Name ="Label120"
                                            Caption ="Display/Enable Duplicate Item"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2790
                                    Top =4440
                                    TabIndex =7
                                    Name ="DropAlertAmt"
                                    ControlSource ="DropAlertAmt"
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
                                            Left =1620
                                            Top =4440
                                            Width =1140
                                            Height =240
                                            Name ="Label71"
                                            Caption ="Drop Alert Amt:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4500
                                    Top =1770
                                    Width =3660
                                    Height =1860
                                    Name ="Box115"
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =4500
                                    Top =3630
                                    Width =3660
                                    Height =1905
                                    Name ="Box116"
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2790
                                    Top =4810
                                    Name ="ShiftCounter"
                                    ControlSource ="ShiftCounter"
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
                                            Left =1770
                                            Top =4810
                                            Width =990
                                            Height =240
                                            Name ="Label51"
                                            Caption ="ShiftCounter:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2790
                                    Top =5170
                                    Name ="LogoutTimer"
                                    ControlSource ="LogoutTimer"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1710
                                            Top =5170
                                            Width =1050
                                            Height =240
                                            Name ="Label83"
                                            Caption ="Logout Timer:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =240
                            Top =1665
                            Width =10440
                            Height =4320
                            Name ="Page88"
                            Caption ="pcCharge"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2075
                                    Top =2025
                                    Name ="pcChargeInstalled"
                                    ControlSource ="pcChargeInstalled"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =420
                                            Top =2025
                                            Width =1590
                                            Height =240
                                            Name ="Label154"
                                            Caption ="pcCharge Installed"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2615
                                    Top =2565
                                    Width =2745
                                    TabIndex =1
                                    Name ="pcChargeDir"
                                    ControlSource ="pcChargeDir"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1500
                                            Top =2565
                                            Width =1050
                                            Height =240
                                            Name ="Label152"
                                            Caption ="pcCharge Dir:"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =3495
                                    Top =2025
                                    Width =2687
                                    Height =418
                                    TabIndex =2
                                    Name ="Frame147"
                                    ControlSource ="pcChargeCommMethod"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2340
                                            Top =2085
                                            Width =1080
                                            Height =240
                                            Name ="Label148"
                                            Caption ="Comm method"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =3555
                                            Top =2113
                                            OptionValue =0
                                            Name ="Option149"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =3785
                                                    Top =2085
                                                    Width =1155
                                                    Height =240
                                                    Name ="Label150"
                                                    Caption ="File Transfer"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =5055
                                            Top =2113
                                            Name ="Option151"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =5285
                                                    Top =2085
                                                    Width =840
                                                    Height =240
                                                    Name ="Label153"
                                                    Caption ="TCP/IP"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =9065
                                    Top =2085
                                    TabIndex =3
                                    Name ="pcChargeIPAddress"
                                    ControlSource ="pcChargeIPAddress"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =8160
                                            Top =2085
                                            Width =840
                                            Height =240
                                            Name ="Label157"
                                            Caption ="IP Address"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5460
                                    Top =2580
                                    Width =1740
                                    Height =240
                                    Name ="Label89"
                                    Caption ="\\\\fs1\\active-charge\\  "
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =9065
                                    Top =2385
                                    TabIndex =4
                                    Name ="pcChargePort"
                                    ControlSource ="pcChargePort"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =8640
                                            Top =2385
                                            Width =360
                                            Height =240
                                            Name ="Label159"
                                            Caption ="Port"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =3
                                    Left =960
                                    Top =3780
                                    Width =1590
                                    Height =240
                                    Name ="Label92"
                                    Caption ="Default to Debit ?"
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =9060
                                    Top =2685
                                    TabIndex =5
                                    Name ="pcChargeEnableSSL"
                                    ControlSource ="pcChargeEnableSSL"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =3
                                            Left =7680
                                            Top =2685
                                            Width =1335
                                            Height =240
                                            Name ="Label163"
                                            Caption ="Enable SSL"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =3120
                                    Top =4080
                                    Width =2370
                                    Height =240
                                    Name ="Label104"
                                    Caption ="Set to '0' if pin pad is not present"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2615
                                    Top =2865
                                    TabIndex =6
                                    Name ="pcChargeProcesser"
                                    ControlSource ="pcChargeProcesser"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =990
                                            Top =2865
                                            Width =1560
                                            Height =240
                                            Name ="Label155"
                                            Caption ="pcCharge Processer:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2615
                                    Top =3165
                                    Width =3150
                                    TabIndex =7
                                    Name ="pcChargeMerchantID"
                                    ControlSource ="pcChargeMerchantID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =855
                                            Top =3165
                                            Width =1695
                                            Height =240
                                            Name ="Label156"
                                            Caption ="pcCharge MerchantID:"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4065
                                    Top =5205
                                    Width =2730
                                    Height =600
                                    Name ="Label109"
                                    Caption ="Leave blank if all credit card processing is done thru OLCC's merchant ID"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =540
                                    Top =1725
                                    Width =7545
                                    Height =195
                                    FontWeight =700
                                    Name ="Label110"
                                    Caption ="PLEASE CONTACT KELLY MYERS BEFORE CHANGING ANY DATA ON THIS SCREEN!"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2610
                                    Top =3480
                                    TabIndex =8
                                    Name ="pcChargeAllowHandEnter"
                                    ControlSource ="pcChargeAllowHandEnter"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =420
                                            Top =3480
                                            Width =2145
                                            Height =240
                                            Name ="Label166"
                                            Caption ="Allow CC to be Hand Entered"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =3105
                                    Top =4860
                                    Width =2745
                                    Height =240
                                    Name ="Label136"
                                    Caption ="Set to '0' if sames as Register Number"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2595
                                    Top =3780
                                    TabIndex =9
                                    Name ="pcChargeDefaultDebit"
                                    ControlSource ="pcChargeDefaultDebit"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =2995
                                            Top =3780
                                            Width =6075
                                            Height =240
                                            Name ="Label93"
                                            Caption ="What should the default response be when the charge key is pressed; Debit or Cre"
                                                "dit"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2600
                                    Top =4080
                                    Width =420
                                    TabIndex =10
                                    Name ="pcChargePinPadComPort"
                                    ControlSource ="pcChargePinPadComPort"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =525
                                            Top =4080
                                            Width =2010
                                            Height =240
                                            Name ="Label103"
                                            Caption ="pcCharge PinPad ComPort:"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =2565
                                    Top =4380
                                    Width =3362
                                    Height =418
                                    TabIndex =11
                                    Name ="Frame138"
                                    ControlSource ="pcChargePinPadTDES"
                                    DefaultValue ="0"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1170
                                            Top =4440
                                            Width =1365
                                            Height =240
                                            Name ="Label139"
                                            Caption ="Pinpad Encryption"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =3045
                                            Top =4468
                                            OptionValue =0
                                            Name ="Option141"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =3275
                                                    Top =4440
                                                    Width =1155
                                                    Height =240
                                                    Name ="Label142"
                                                    Caption ="Master Session"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =4665
                                            Top =4468
                                            Name ="Option143"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =4895
                                                    Top =4440
                                                    Width =840
                                                    Height =240
                                                    Name ="Label144"
                                                    Caption ="Triple DES"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =2585
                                    Top =4860
                                    Width =420
                                    TabIndex =12
                                    Name ="pcChargeUserNo"
                                    ControlSource ="pcChargeUserNo"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =525
                                            Top =4860
                                            Width =2010
                                            Height =240
                                            Name ="Label135"
                                            Caption ="pcCharge User Number"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2585
                                    Top =5215
                                    TabIndex =13
                                    Name ="pcChargeProcesserRelated"
                                    ControlSource ="pcChargeProcesserRelated"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =375
                                            Top =5205
                                            Width =2160
                                            Height =240
                                            Name ="Label106"
                                            Caption ="Related pcCharge Processer:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =247
                                    Left =7380
                                    Top =1965
                                    Width =3300
                                    Height =1080
                                    Name ="Box164"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2585
                                    Top =5515
                                    Width =1425
                                    TabIndex =14
                                    Name ="pcChargeMerchantIDRelated"
                                    ControlSource ="pcChargeMerchantIDRelated"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =240
                                            Top =5505
                                            Width =2295
                                            Height =240
                                            Name ="Label108"
                                            Caption ="Related pcCharge MerchantID:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =240
                            Top =1665
                            Width =10440
                            Height =4320
                            Name ="Page121"
                            Caption ="DVR Interface"
                            Begin
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1980
                                    Top =1785
                                    Name ="DVRRemoteHost"
                                    ControlSource ="DVRRemoteHost"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =540
                                            Top =1785
                                            Width =1395
                                            Height =240
                                            Name ="Label123"
                                            Caption ="DVR RemoteHost:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1980
                                    Top =2085
                                    Name ="DVRRemotePort"
                                    ControlSource ="DVRRemotePort"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =540
                                            Top =2085
                                            Width =1350
                                            Height =240
                                            Name ="Label124"
                                            Caption ="DVR RemotePort:"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =420
                                    Top =2565
                                    Width =3426
                                    Height =943
                                    Name ="Frame127"
                                    ControlSource ="DVRProtocol"
                                    DefaultValue ="0"

                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            Left =540
                                            Top =2445
                                            Width =2055
                                            Height =240
                                            Name ="Label128"
                                            Caption ="DVR Protocol (usually UDP)"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =1380
                                            Top =2833
                                            OptionValue =0
                                            Name ="Option130"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =1610
                                                    Top =2805
                                                    Width =390
                                                    Height =240
                                                    Name ="Label131"
                                                    Caption ="TCP"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =1380
                                            Top =3163
                                            OptionValue =1
                                            Name ="Option132"

                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =1610
                                                    Top =3135
                                                    Width =420
                                                    Height =240
                                                    Name ="Label133"
                                                    Caption ="UDP"
                                                End
                                            End
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2175
                    Top =600
                    TabIndex =3
                    Name ="ComputerName"
                    ControlSource ="ComputerName"
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
                            Left =900
                            Top =600
                            Width =1215
                            Height =240
                            Name ="Label50"
                            Caption ="ComputerName:"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =6120
                    Width =10800
                    Height =420
                    Name ="Box13"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =1335
                    Top =6180
                    Width =1890
                    Height =255
                    TabIndex =6
                    BackColor =12632256
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =780
                            Top =6180
                            Width =495
                            Height =240
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =7410
                    Top =6240
                    Width =1890
                    Height =255
                    TabIndex =7
                    BackColor =12632256
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="General Date"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =6720
                            Top =6240
                            Width =690
                            Height =240
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5280
                    Top =240
                    TabIndex =2
                    Name ="chkActive"
                    ControlSource ="XfrSmode"
                    DefaultValue ="True"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =240
                    LayoutCachedWidth =5540
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4380
                            Top =240
                            Width =810
                            Height =240
                            Name ="Label170"
                            Caption ="Active"
                            LayoutCachedLeft =4380
                            LayoutCachedTop =240
                            LayoutCachedWidth =5190
                            LayoutCachedHeight =480
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6180
                    Top =300
                    Width =3255
                    Height =315
                    TabIndex =8
                    Name ="cmdCashRegisterSystemStatus"
                    Caption ="Check Online Status of Active Workstations"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =300
                    LayoutCachedWidth =9435
                    LayoutCachedHeight =615
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =1005
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =4620
                    Top =60
                    Width =666
                    Height =606
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =4658
                            Top =660
                            Width =555
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =3180
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =1
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3247
                            Top =660
                            Width =420
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1740
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =2
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1807
                            Top =660
                            Width =450
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =8820
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =3
                    Name ="cmdfrmClose"
                    Caption ="Command64"
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

                    LayoutCachedLeft =8820
                    LayoutCachedTop =120
                    LayoutCachedWidth =9486
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =8820
                            Top =720
                            Width =885
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label78"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =8820
                            LayoutCachedTop =720
                            LayoutCachedWidth =9705
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =360
                    Top =60
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =368
                            Top =660
                            Width =705
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =60
                    Width =10800
                    Height =960
                    Name ="Box82"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5880
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

                    LayoutCachedLeft =5880
                    LayoutCachedTop =120
                    LayoutCachedWidth =6546
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =5880
                            Top =720
                            Width =1095
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label167"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5880
                            LayoutCachedTop =720
                            LayoutCachedWidth =6975
                            LayoutCachedHeight =1005
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6960
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

                    LayoutCachedLeft =6960
                    LayoutCachedTop =120
                    LayoutCachedWidth =7626
                    LayoutCachedHeight =726
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =6960
                            Top =720
                            Width =915
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6960
                            LayoutCachedTop =720
                            LayoutCachedWidth =7875
                            LayoutCachedHeight =1005
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CashRegisterSystem.cls"
