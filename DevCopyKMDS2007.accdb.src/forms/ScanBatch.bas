Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14310
    DatasheetFontHeight =11
    ItemSuffix =155
    Left =2175
    Top =1545
    Right =16845
    Bottom =8970
    RecSrcDt = Begin
        0xb108895547fee340
    End
    RecordSource ="qryScanBatchLineCounts"
    Caption ="Scan Batch at 1. Storeshelf"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =1740
            Name ="FormHeader"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Width =3060
                    Height =315
                    ColumnOrder =0
                    Name ="txtScanBatchHeaderId"
                    DefaultValue ="3"

                    LayoutCachedLeft =3660
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Width =3600
                            Height =315
                            Name ="Label45"
                            Caption ="Scan Batch Header Id"
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3660
                    Top =435
                    Width =3060
                    Height =315
                    ColumnOrder =2
                    TabIndex =1
                    Name ="txtCashierID"
                    DefaultValue ="1"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =435
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =435
                            Width =3600
                            Height =315
                            Name ="Label47"
                            Caption ="CashierID"
                            LayoutCachedTop =435
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =750
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3660
                    Top =870
                    Width =3060
                    Height =315
                    ColumnOrder =3
                    TabIndex =2
                    Name ="txtBatchHeaderDescription"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =870
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1185
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =870
                            Width =3600
                            Height =315
                            Name ="Label49"
                            Caption ="Batch Header Description"
                            LayoutCachedTop =870
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =1185
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3660
                    Top =1305
                    Width =3060
                    Height =315
                    ColumnOrder =4
                    TabIndex =3
                    Name ="txtBatchHeaderSetupDate"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =1305
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1620
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1305
                            Width =3600
                            Height =315
                            Name ="Label51"
                            Caption ="BatchHeaderSetupDate"
                            LayoutCachedTop =1305
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7620
                    Top =1200
                    Width =2220
                    Height =315
                    ColumnOrder =5
                    TabIndex =4
                    Name ="txtBatchHeaderChangeDate"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1515
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5160
                            Top =1200
                            Width =2400
                            Height =315
                            Name ="Label53"
                            Caption ="Batch Header Change Date"
                            LayoutCachedLeft =5160
                            LayoutCachedTop =1200
                            LayoutCachedWidth =7560
                            LayoutCachedHeight =1515
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8460
                    Width =3060
                    Height =315
                    ColumnOrder =6
                    TabIndex =5
                    Name ="txtItemID"

                    LayoutCachedLeft =8460
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7380
                            Width =1020
                            Height =315
                            Name ="Label37"
                            Caption ="ItemID"
                            LayoutCachedLeft =7380
                            LayoutCachedWidth =8400
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =375
                    Width =3060
                    Height =315
                    ColumnOrder =7
                    TabIndex =6
                    Name ="txtOnShelf"
                    Format ="General Number"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =375
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5580
                            Top =360
                            Width =1140
                            Height =315
                            Name ="Label41"
                            Caption ="Onshelf"
                            LayoutCachedLeft =5580
                            LayoutCachedTop =360
                            LayoutCachedWidth =6720
                            LayoutCachedHeight =675
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =795
                    Width =3060
                    Height =315
                    ColumnOrder =8
                    TabIndex =7
                    Name ="txtInputMethod"
                    Format ="General Number"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =795
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5280
                            Top =780
                            Width =1440
                            Height =315
                            Name ="Label55"
                            Caption ="Input Method"
                            LayoutCachedLeft =5280
                            LayoutCachedTop =780
                            LayoutCachedWidth =6720
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =11100
                    Top =180
                    Width =2880
                    Height =1140
                    AutoActivate =1
                    TabIndex =8
                    Name ="WMP1"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000522af56b4a39d311b15300c0 ,
                        0x4f79faa6000000000000000000000000707805ea9cb7cc0107000000c0010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000080010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff01000000522af56b4a39d311b15300c0 ,
                        0x4f79faa600000000000000000000000000507d6c752ccd010500000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000001000000ec000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff020000000300000004000000feffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x00000000522af56b4a39d311b15300c04f79faa60003000008003a0000004300 ,
                        0x3a005c004b004d00440053005c0053006f0075006e00640073005c004e006f00 ,
                        0x74004f006e00460069006c0065002e0057006100760000000500000000000000 ,
                        0xf03f030000000000050000000000000000000800020000000000030001000000 ,
                        0x0b00ffff0300000000000b00ffff08000200000000000300320000000b000000 ,
                        0x08000a000000660075006c006c0000000b0000000b0000000b00ffff0b00ffff ,
                        0x0b00000008000200000000000800020000000000080002000000000008000200 ,
                        0x000000000b000000ce1800005219000000000000f03f03000000000005000000 ,
                        0x00000000000008000200000000000300010000000b00ffff0300000000000b00 ,
                        0xffff08000200000000000300320000000b00000008000a000000660075006c00 ,
                        0x6c0000000b0000000b0000000b00ffff0b00ffff0b0000000800020000000000 ,
                        0x0800020000000000080002000000000008000200000000000b000000ce180000 ,
                        0x52190000000000000b000000ce18000052190000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="Windows Media Player"
                    Class ="WMPlayer.OCX.7"

                    LayoutCachedLeft =11100
                    LayoutCachedTop =180
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1320
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6600
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9300
                    Top =2040
                    Width =1395
                    Height =630
                    FontSize =24
                    TabIndex =21
                    BackColor =62207
                    Name ="txtPriceLast"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =9300
                    LayoutCachedTop =2040
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =2670
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =204
                            TextAlign =2
                            Left =9540
                            Top =1740
                            Width =1275
                            Height =270
                            FontSize =9
                            Name ="lblPriceForInquiry"
                            Caption ="Price"
                            LayoutCachedLeft =9540
                            LayoutCachedTop =1740
                            LayoutCachedWidth =10815
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =180
                    Top =840
                    Width =1140
                    FontSize =8
                    TabIndex =15
                    Name ="txtLocation"

                    LayoutCachedLeft =180
                    LayoutCachedTop =840
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1320
                    Top =240
                    Width =3420
                    Height =630
                    FontSize =24
                    Name ="txtCode"
                    StatusBarText ="Double Click Here to see more options.."
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    ControlTipText ="Double Click Here to see more options.."
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =1320
                    LayoutCachedTop =240
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =870
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =240
                            Width =1020
                            Height =315
                            Name ="Label1"
                            Caption ="Scan Code"
                            LayoutCachedLeft =180
                            LayoutCachedTop =240
                            LayoutCachedWidth =1200
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4860
                    Top =240
                    Width =7395
                    Height =630
                    FontSize =20
                    TabIndex =3
                    Name ="txtDescription"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =240
                    LayoutCachedWidth =12255
                    LayoutCachedHeight =870
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =120
                    Width =14220
                    Height =6480
                    Name ="Box67"
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =6600
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9060
                    Top =960
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="txtScanCounts"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =960
                    LayoutCachedWidth =9990
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =3
                            Left =8040
                            Top =960
                            Width =930
                            Height =600
                            FontSize =9
                            Name ="Label69"
                            Caption ="Scans  ALL "
                            LayoutCachedLeft =8040
                            LayoutCachedTop =960
                            LayoutCachedWidth =8970
                            LayoutCachedHeight =1560
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5760
                    Top =960
                    Width =1995
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="txtFirstScanDate"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =960
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =3
                            Left =4800
                            Top =960
                            Width =870
                            Height =270
                            FontSize =9
                            Name ="Label71"
                            Caption ="First Scan:"
                            LayoutCachedLeft =4800
                            LayoutCachedTop =960
                            LayoutCachedWidth =5670
                            LayoutCachedHeight =1230
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5760
                    Top =1320
                    Width =1995
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="txtLastScanDate"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =3
                            Left =4860
                            Top =1320
                            Width =870
                            Height =270
                            FontSize =9
                            Name ="Label73"
                            Caption ="Last Scan:"
                            LayoutCachedLeft =4860
                            LayoutCachedTop =1320
                            LayoutCachedWidth =5730
                            LayoutCachedHeight =1590
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =300
                    Top =2040
                    Width =2835
                    Height =630
                    FontSize =24
                    TabIndex =7
                    BackColor =62207
                    Name ="txtCodeLast"
                    StatusBarText ="Double Click Here to see more options.."
                    ControlTipText ="Double Click Here to see more options.."
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =300
                    LayoutCachedTop =2040
                    LayoutCachedWidth =3135
                    LayoutCachedHeight =2670
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3240
                    Top =2040
                    Width =5985
                    Height =630
                    FontSize =20
                    TabIndex =8
                    BackColor =62207
                    Name ="txtDescriptionLast"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2040
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =2670
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10800
                    Top =2040
                    Width =1035
                    Height =630
                    FontSize =24
                    TabIndex =9
                    BackColor =62207
                    Name ="txtCountLast"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =10800
                    LayoutCachedTop =2040
                    LayoutCachedWidth =11835
                    LayoutCachedHeight =2670
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =2
                            Left =10860
                            Top =1740
                            Width =1080
                            Height =270
                            FontSize =9
                            Name ="lblCountLast"
                            Caption ="Count (1)"
                            LayoutCachedLeft =10860
                            LayoutCachedTop =1740
                            LayoutCachedWidth =11940
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11940
                    Top =2040
                    Width =1035
                    Height =630
                    FontSize =24
                    TabIndex =10
                    BackColor =62207
                    Name ="txtOnShelfLast"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =11940
                    LayoutCachedTop =2040
                    LayoutCachedWidth =12975
                    LayoutCachedHeight =2670
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =2
                            Left =11940
                            Top =1740
                            Width =960
                            Height =270
                            FontSize =9
                            Name ="Label85"
                            Caption ="On Shelf"
                            LayoutCachedLeft =11940
                            LayoutCachedTop =1740
                            LayoutCachedWidth =12900
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13080
                    Top =2040
                    Width =1035
                    Height =630
                    FontSize =24
                    TabIndex =12
                    BackColor =65280
                    Name ="txtVar"

                    LayoutCachedLeft =13080
                    LayoutCachedTop =2040
                    LayoutCachedWidth =14115
                    LayoutCachedHeight =2670
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =2
                            Left =13080
                            Top =1740
                            Width =1020
                            Height =270
                            FontSize =9
                            Name ="Label86"
                            Caption ="Variance"
                            LayoutCachedLeft =13080
                            LayoutCachedTop =1740
                            LayoutCachedWidth =14100
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =1680
                    Width =14040
                    Height =1080
                    Name ="Box90"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1680
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =2760
                End
                Begin Label
                    OverlapFlags =255
                    Left =360
                    Top =1680
                    Width =1320
                    Height =315
                    Name ="Label91"
                    Caption ="Last Scan Info"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1680
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1995
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =2880
                    Width =13440
                    Height =3600
                    Name ="Box92"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2880
                    LayoutCachedWidth =13620
                    LayoutCachedHeight =6480
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1740
                    Top =1740
                    Width =1695
                    FontSize =9
                    TabIndex =13
                    BackColor =62207
                    Name ="txtItemIDLast"
                    StatusBarText ="Double Click Here to see more options.."
                    ControlTipText ="Double Click Here to see more options.."

                    LayoutCachedLeft =1740
                    LayoutCachedTop =1740
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =180
                    Top =1080
                    Width =1140
                    FontSize =8
                    TabIndex =14
                    Name ="ItemID"
                    ControlSource ="ItemID"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =180
                    Top =1320
                    FontSize =8
                    TabIndex =16
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1560
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12360
                    Top =240
                    Width =1680
                    Height =630
                    FontSize =24
                    TabIndex =17
                    Name ="txtPrice"
                    Format ="Standard"
                    StatusBarText ="Double Click Here to see more options.."
                    ControlTipText ="Double Click Here to see more options.."
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =12360
                    LayoutCachedTop =240
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =870
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1290
                    Top =960
                    Width =3420
                    Height =630
                    FontSize =24
                    TabIndex =2
                    Name ="txtCount"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =1290
                    LayoutCachedTop =960
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =1590
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000a7da4e00000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =540
                            Top =960
                            Width =630
                            Height =315
                            Name ="Label3"
                            Caption ="Count"
                            LayoutCachedLeft =540
                            LayoutCachedTop =960
                            LayoutCachedWidth =1170
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8520
                    Top =2940
                    Width =855
                    FontSize =9
                    TabIndex =18
                    Name ="txtOnShelfDate"
                    Format ="Medium Time"

                    LayoutCachedLeft =8520
                    LayoutCachedTop =2940
                    LayoutCachedWidth =9375
                    LayoutCachedHeight =3180
                End
                Begin OptionButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =12540
                    Top =990
                    TabIndex =19
                    Name ="optVoice"
                    DefaultValue ="False"

                    LayoutCachedLeft =12540
                    LayoutCachedTop =990
                    LayoutCachedWidth =12800
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =12765
                            Top =960
                            Width =1305
                            Height =315
                            Name ="Label114"
                            Caption ="Google Voice"
                            LayoutCachedLeft =12765
                            LayoutCachedTop =960
                            LayoutCachedWidth =14070
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin OptionButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10140
                    Top =990
                    TabIndex =20
                    Name ="optPriceCheckMode"
                    DefaultValue ="False"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =990
                    LayoutCachedWidth =10400
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =10365
                            Top =960
                            Width =1695
                            Height =315
                            Name ="Label116"
                            Caption ="Price check mode "
                            LayoutCachedLeft =10365
                            LayoutCachedTop =960
                            LayoutCachedWidth =12060
                            LayoutCachedHeight =1275
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =86
                    Left =1740
                    Top =3780
                    Width =1680
                    Height =720
                    TabIndex =22
                    Name ="cmd65VarReport"
                    Caption ="&Variance Report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print Variance report by item code"
                    UnicodeAccessKey =86

                    LayoutCachedLeft =1740
                    LayoutCachedTop =3780
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =4500
                    Overlaps =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =1560
                            Top =3420
                            Width =2055
                            Height =240
                            FontSize =9
                            Name ="Label147"
                            Caption ="Reports"
                            LayoutCachedLeft =1560
                            LayoutCachedTop =3420
                            LayoutCachedWidth =3615
                            LayoutCachedHeight =3660
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =80
                    Left =1755
                    Top =5340
                    Width =1680
                    Height =720
                    TabIndex =23
                    Name ="cmdBatchScanReport"
                    Caption ="&Print Scan Qty Barcodes"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Scan Codes for Quantity Entries"
                    UnicodeAccessKey =80
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1755
                    LayoutCachedTop =5340
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =6060
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =69
                    Left =4455
                    Top =3780
                    Width =1680
                    Height =720
                    TabIndex =24
                    Name ="cmdEraseBatch"
                    Caption ="&Erase this batch!"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =69
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4455
                    LayoutCachedTop =3780
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =4500
                    Overlaps =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =4260
                            Top =3420
                            Width =1995
                            Height =240
                            FontSize =9
                            Name ="Label148"
                            Caption ="Records"
                            LayoutCachedLeft =4260
                            LayoutCachedTop =3420
                            LayoutCachedWidth =6255
                            LayoutCachedHeight =3660
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =65
                    Left =6720
                    Top =3780
                    Width =1680
                    Height =720
                    TabIndex =25
                    Name ="cmdAppendToAudit"
                    Caption ="Move to D-7 &Audit"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =65
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =3780
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =4500
                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =77
                    Left =13680
                    Top =5160
                    Width =606
                    Height =606
                    FontSize =8
                    TabIndex =26
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13680
                    LayoutCachedTop =5160
                    LayoutCachedWidth =14286
                    LayoutCachedHeight =5766
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =13740
                            Top =5880
                            Width =510
                            Height =525
                            FontSize =8
                            Name ="Label65"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =13740
                            LayoutCachedTop =5880
                            LayoutCachedWidth =14250
                            LayoutCachedHeight =6405
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6720
                    Top =4560
                    Width =1680
                    Height =720
                    TabIndex =27
                    Name ="cmdCreatePO"
                    Caption ="Create Purchase Order"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =4560
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =5280
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =3840
                    Top =240
                    Width =7395
                    Height =2520
                    TabIndex =1
                    BackColor =5167783
                    Name ="lstLocation"
                    RowSourceType ="Value List"
                    RowSource ="Select a Location;1. Storeshelf;2. Backroom;3. Overflow;4. Display"
                    ColumnWidths ="1440"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =240
                    LayoutCachedWidth =11235
                    LayoutCachedHeight =2760
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =180
                    Top =240
                    Width =3900
                    Height =1980
                    TabIndex =28
                    BoundColumn =1
                    BackColor =11193849
                    Name ="lstVendor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Name, ID FROM Vendor WHERE (((Vendor.ID)>0)) ORDER BY Vendor.Name; "
                    ColumnWidths ="2880;720"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =240
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =2220
                End
                Begin OptionButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =10140
                    Top =1350
                    TabIndex =29
                    Name ="optAutoEnter1Qty"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If \"2. Backroom\" is selected then Case Quantity is entered"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =1350
                    LayoutCachedWidth =10400
                    LayoutCachedHeight =1590
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =10365
                            Top =1320
                            Width =2175
                            Height =300
                            Name ="Label125"
                            Caption ="Auto Enter 1 Quantity "
                            ControlTipText ="If \"2. Backroom\" is selected then Case Quantity is entered"
                            LayoutCachedLeft =10365
                            LayoutCachedTop =1320
                            LayoutCachedWidth =12540
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =83
                    Left =1755
                    Top =4560
                    Width =1680
                    Height =720
                    TabIndex =30
                    Name ="cmdBatchReport"
                    Caption ="&Scan Log Report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print all entries in this batch in the order they were entered."
                    UnicodeAccessKey =83

                    LayoutCachedLeft =1755
                    LayoutCachedTop =4560
                    LayoutCachedWidth =3435
                    LayoutCachedHeight =5280
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6720
                    Top =5340
                    Width =1680
                    Height =720
                    TabIndex =31
                    Name ="cmdCreateClubOrder"
                    Caption ="Create Club Order"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6720
                    LayoutCachedTop =5340
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =6060
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =180
                    Top =240
                    Width =3840
                    Height =1800
                    TabIndex =32
                    BoundColumn =1
                    BackColor =15590879
                    Name ="lstClubs"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Name, ID FROM Club WHERE (((ID)>0)) ORDER BY Name; "
                    ColumnWidths ="2880;720"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =240
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =2040
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7920
                    Top =2940
                    Width =555
                    FontSize =9
                    TabIndex =33
                    Name ="txtTotalQty"

                    LayoutCachedLeft =7920
                    LayoutCachedTop =2940
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =3180
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =3
                    Left =6600
                    Top =2940
                    Width =1200
                    Height =240
                    FontSize =9
                    Name ="Label134"
                    Caption ="Total Count:"
                    LayoutCachedLeft =6600
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =3180
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9060
                    Top =1320
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =11
                    Name ="txtScanQ"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =1320
                    LayoutCachedWidth =9990
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =3
                            Left =8100
                            Top =1320
                            Width =855
                            Height =225
                            FontSize =9
                            Name ="lblScanQ"
                            Caption ="1402"
                            LayoutCachedLeft =8100
                            LayoutCachedTop =1320
                            LayoutCachedWidth =8955
                            LayoutCachedHeight =1545
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =82
                    Left =4455
                    Top =5340
                    Width =1680
                    Height =720
                    TabIndex =34
                    Name ="cmdRestoreBatch"
                    Caption ="&Restore Batch"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =82
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4455
                    LayoutCachedTop =5340
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =6060
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    AccessKey =67
                    Left =4455
                    Top =4560
                    Width =1680
                    Height =720
                    TabIndex =35
                    Name ="cmdChangeLocations"
                    Caption ="&Change Locations"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4455
                    LayoutCachedTop =4560
                    LayoutCachedWidth =6135
                    LayoutCachedHeight =5280
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =8520
                    Top =3780
                    Width =1680
                    Height =720
                    TabIndex =36
                    Name ="cmdCreateTransfer"
                    Caption ="Create Store Transfer"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =3780
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =4500
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =8520
                    Top =5340
                    Width =1680
                    Height =720
                    TabIndex =37
                    Name ="cmdCreateClaim"
                    Caption ="Create Claim"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =5340
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =6060
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =480
                    Top =240
                    Width =7395
                    Height =2520
                    TabIndex =38
                    BackColor =13816815
                    Name ="lstAgents"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Agents.AgentNumber, Agents.LastName, Agents.ActiveForTransfer FROM Agents"
                        " WHERE (((Agents.LastName) Not Like \"*Claim*\") AND ((Agents.ActiveForTransfer)"
                        "=True)) ORDER BY Agents.AgentNumber; "
                    ColumnWidths ="720;2880"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =480
                    LayoutCachedTop =240
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =2760
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =180
                    Top =420
                    Width =2940
                    Height =840
                    TabIndex =39
                    BoundColumn =1
                    Name ="lstClaim"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Agents.AgentNumber, Agents.LastName, Agents.ActiveForTransfer FROM Agents"
                        " WHERE (((Agents.LastName) Like \"*Claim*\") AND ((Agents.ActiveForTransfer)=Tru"
                        "e)) ORDER BY Agents.AgentNumber; "
                    ColumnWidths ="720;2880"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1260
                    BackThemeColorIndex =9
                    BackTint =60.0
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =6675
                    Top =3660
                    Width =3615
                    Height =2520
                    Name ="Box144"
                    LayoutCachedLeft =6675
                    LayoutCachedTop =3660
                    LayoutCachedWidth =10290
                    LayoutCachedHeight =6180
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =4275
                    Top =3660
                    Width =1980
                    Height =2520
                    Name ="Box145"
                    LayoutCachedLeft =4275
                    LayoutCachedTop =3660
                    LayoutCachedWidth =6255
                    LayoutCachedHeight =6180
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =1575
                    Top =3660
                    Width =2040
                    Height =2520
                    Name ="Box146"
                    LayoutCachedLeft =1575
                    LayoutCachedTop =3660
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =6180
                End
                Begin Label
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =2
                    Left =6660
                    Top =3420
                    Width =3615
                    Height =240
                    FontSize =9
                    Name ="Label149"
                    Caption ="Move scan data to one of the following."
                    LayoutCachedLeft =6660
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10275
                    LayoutCachedHeight =3660
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =180
                    Top =3150
                    Width =13380
                    Height =3240
                    TabIndex =40
                    Name ="ScanBatchSub"
                    SourceObject ="Form.ScanBatchSub"

                    LayoutCachedLeft =180
                    LayoutCachedTop =3150
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =6390
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =240
                            Top =2850
                            Width =2490
                            Height =315
                            Name ="ScanBatchLineCountsSub2 Label"
                            Caption ="Previous Entries"
                            EventProcPrefix ="ScanBatchLineCountsSub2_Label"
                            LayoutCachedLeft =240
                            LayoutCachedTop =2850
                            LayoutCachedWidth =2730
                            LayoutCachedHeight =3165
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =8520
                    Top =4560
                    Width =1680
                    Height =720
                    TabIndex =41
                    Name ="cmdComparePO"
                    Caption ="Compare to Purchase Order"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8520
                    LayoutCachedTop =4560
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =5280
                    Overlaps =1
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =240
                    Top =300
                    Width =7500
                    Height =1800
                    TabIndex =42
                    BoundColumn =1
                    BackColor =15590879
                    Name ="lstPurchaseOrders"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PurchaseOrder.ID, Vendor.Name, qryPurchaseOrderLineTotals.Quantity, qryPu"
                        "rchaseOrderLineTotals.Dollars FROM (PurchaseOrder INNER JOIN Vendor ON PurchaseO"
                        "rder.VendorID = Vendor.ID) INNER JOIN qryPurchaseOrderLineTotals ON PurchaseOrde"
                        "r.ID = qryPurchaseOrderLineTotals.PurchaseOrderID; "
                    ColumnWidths ="2880;720"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =300
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =2100
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =10380
                    Top =4740
                    Height =315
                    TabIndex =43
                    Name ="txtOrderID"

                    LayoutCachedLeft =10380
                    LayoutCachedTop =4740
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =5055
                End
            End
        End
        Begin FormFooter
            Height =750
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =14220
                    Height =630
                    FontSize =24
                    BackColor =65535
                    Name ="txtStatus"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =690
                End
            End
        End
    End
End
CodeBehindForm
' See "ScanBatch.cls"
