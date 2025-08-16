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
    Width =12600
    DatasheetFontHeight =11
    ItemSuffix =112
    Left =-3525
    Top =780
    Right =9165
    Bottom =8145
    RecSrcDt = Begin
        0x291212cc67f6e340
    End
    RecordSource ="qryScanBatchLineCounts"
    Caption ="Scan Batch at 1. Store Shelf"
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =1620
            Name ="FormHeader"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Width =3060
                    Height =315
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8460
                    Width =3060
                    Height =315
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6600
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =180
                    Top =840
                    Width =1140
                    FontSize =8
                    TabIndex =17
                    Name ="txtLocation"

                    LayoutCachedLeft =180
                    LayoutCachedTop =840
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =1080
                End
                Begin Subform
                    Visible = NotDefault
                    CanGrow = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =840
                    Top =4020
                    Width =10020
                    Height =2040
                    TabIndex =14
                    Name ="ScanBatchLineCountsSub"
                    SourceObject ="Form.ScanBatchLineCountsSub"
                    LinkChildFields ="ItemID"
                    LinkMasterFields ="ItemID"

                    LayoutCachedLeft =840
                    LayoutCachedTop =4020
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =6060
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =3060
                            Top =2940
                            Width =1770
                            Height =315
                            Name ="ScanBatchLineCountsSubLabel"
                            Caption ="Scans for this item"
                            LayoutCachedLeft =3060
                            LayoutCachedTop =2940
                            LayoutCachedWidth =4830
                            LayoutCachedHeight =3255
                        End
                    End
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
                    OnDblClick ="[Event Procedure]"
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
                    Width =5820
                    Height =630
                    FontSize =20
                    TabIndex =4
                    Name ="txtDescription"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =240
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =870
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =105
                    Top =120
                    Width =12495
                    Height =6480
                    Name ="Box67"
                    LayoutCachedLeft =105
                    LayoutCachedTop =120
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =6600
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11550
                    Top =960
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =5
                    Name ="txtScanCounts"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11550
                    LayoutCachedTop =960
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =3
                            Left =10500
                            Top =960
                            Width =930
                            Height =600
                            FontSize =9
                            Name ="Label69"
                            Caption ="Scans  ALL "
                            LayoutCachedLeft =10500
                            LayoutCachedTop =960
                            LayoutCachedWidth =11430
                            LayoutCachedHeight =1560
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7545
                    Top =960
                    Width =2880
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="txtFirstScanDate"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =960
                    LayoutCachedWidth =10425
                    LayoutCachedHeight =1215
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =204
                            TextAlign =3
                            Left =6600
                            Top =960
                            Width =870
                            Height =270
                            FontSize =9
                            Name ="Label71"
                            Caption ="First Scan:"
                            LayoutCachedLeft =6600
                            LayoutCachedTop =960
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =1230
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7545
                    Top =1320
                    Width =2880
                    Height =255
                    FontSize =9
                    TabIndex =7
                    Name ="txtLastScanDate"

                    LayoutCachedLeft =7545
                    LayoutCachedTop =1320
                    LayoutCachedWidth =10425
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =204
                            TextAlign =3
                            Left =6600
                            Top =1320
                            Width =870
                            Height =270
                            FontSize =9
                            Name ="Label73"
                            Caption ="Last Scan:"
                            LayoutCachedLeft =6600
                            LayoutCachedTop =1320
                            LayoutCachedWidth =7470
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
                    Width =3120
                    Height =630
                    FontSize =24
                    TabIndex =8
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
                    LayoutCachedWidth =3420
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
                    Left =3495
                    Top =2040
                    Width =5985
                    Height =630
                    FontSize =20
                    TabIndex =9
                    BackColor =62207
                    Name ="txtDescriptionLast"

                    LayoutCachedLeft =3495
                    LayoutCachedTop =2040
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =2670
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9570
                    Top =2040
                    Width =870
                    Height =630
                    FontSize =24
                    TabIndex =10
                    BackColor =62207
                    Name ="txtCountLast"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =9570
                    LayoutCachedTop =2040
                    LayoutCachedWidth =10440
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
                            Left =9570
                            Top =1740
                            Width =780
                            Height =270
                            FontSize =9
                            Name ="lblCountLast"
                            Caption ="Count (2)"
                            LayoutCachedLeft =9570
                            LayoutCachedTop =1740
                            LayoutCachedWidth =10350
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
                    Left =10500
                    Top =2040
                    Width =900
                    Height =630
                    FontSize =24
                    TabIndex =11
                    BackColor =62207
                    Name ="txtOnShelfLast"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =10500
                    LayoutCachedTop =2040
                    LayoutCachedWidth =11400
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
                            Left =10500
                            Top =1740
                            Width =780
                            Height =270
                            FontSize =9
                            Name ="Label85"
                            Caption ="On Shelf"
                            LayoutCachedLeft =10500
                            LayoutCachedTop =1740
                            LayoutCachedWidth =11280
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11550
                    Top =1320
                    Width =930
                    Height =255
                    FontSize =9
                    TabIndex =12
                    Name ="txtScanQ"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =11550
                    LayoutCachedTop =1320
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =1575
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =204
                            TextAlign =3
                            Left =10500
                            Top =1320
                            Width =930
                            Height =270
                            FontSize =9
                            Name ="lblScanQ"
                            Caption ="60952"
                            LayoutCachedLeft =10500
                            LayoutCachedTop =1320
                            LayoutCachedWidth =11430
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
                    Left =11460
                    Top =2040
                    Width =900
                    Height =630
                    FontSize =24
                    TabIndex =13
                    BackColor =255
                    Name ="txtVar"

                    LayoutCachedLeft =11460
                    LayoutCachedTop =2040
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =2670
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =204
                            TextAlign =2
                            Left =11460
                            Top =1740
                            Width =780
                            Height =270
                            FontSize =9
                            Name ="Label86"
                            Caption ="Variance"
                            LayoutCachedLeft =11460
                            LayoutCachedTop =1740
                            LayoutCachedWidth =12240
                            LayoutCachedHeight =2010
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =225
                    Top =1680
                    Width =12255
                    Height =1080
                    Name ="Box90"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1680
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =2760
                End
                Begin Label
                    OverlapFlags =247
                    Left =360
                    Top =1740
                    Width =1320
                    Height =315
                    Name ="Label91"
                    Caption ="Last Scan Info"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1740
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =2055
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =330
                    Top =2880
                    Width =12030
                    Height =3600
                    Name ="Box92"
                    LayoutCachedLeft =330
                    LayoutCachedTop =2880
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =6480
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    TextFontCharSet =204
                    IMESentenceMode =3
                    ColumnCount =8
                    Left =540
                    Top =3240
                    Width =11640
                    Height =3120
                    TabIndex =3
                    Name ="lstEntries"
                    RowSourceType ="Value List"
                    RowSource ="UPC Code;Item Id;Description;Count;Shelf;Location;InputMethod;1402;1402;n/a;10;5"
                        "1;2. Back Shelf;Normal entry;088110931032;60952;n/a;3;92;1. Store Shelf;Normal e"
                        "ntry;088110931032;60952;n/a;12;92;4. Display;Normal entry"
                    ColumnWidths ="1440;1080;2880;720;720;720;1440"

                    LayoutCachedLeft =540
                    LayoutCachedTop =3240
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =6360
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =480
                            Top =2940
                            Width =1680
                            Height =315
                            Name ="Label93"
                            Caption ="Scans as entered:"
                            LayoutCachedLeft =480
                            LayoutCachedTop =2940
                            LayoutCachedWidth =2160
                            LayoutCachedHeight =3255
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1740
                    Top =1740
                    Width =1200
                    FontSize =9
                    TabIndex =15
                    BackColor =62207
                    Name ="txtItemIDLast"
                    StatusBarText ="Double Click Here to see more options.."
                    ControlTipText ="Double Click Here to see more options.."

                    LayoutCachedLeft =1740
                    LayoutCachedTop =1740
                    LayoutCachedWidth =2940
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
                    TabIndex =16
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
                    TabIndex =18
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1560
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3900
                    Top =240
                    Width =1920
                    Height =2460
                    TabIndex =1
                    BackColor =5167783
                    Name ="lstLocation"
                    RowSourceType ="Value List"
                    RowSource ="Select a Location;1. Store Shelf;2. Back Shelf;3. Overflow;4. Display;5. Not Ass"
                        "igned;6. Not Assigned;7. Not Assigned;8. Not Assigned"
                    ColumnWidths ="1440"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =240
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =2700
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10800
                    Top =240
                    Width =1680
                    Height =630
                    FontSize =24
                    TabIndex =19
                    Name ="txtPrice"
                    StatusBarText ="Double Click Here to see more options.."
                    ControlTipText ="Double Click Here to see more options.."
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x00000000a7da4e00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =10800
                    LayoutCachedTop =240
                    LayoutCachedWidth =12480
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
                            OverlapFlags =247
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
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10500
                    Top =2940
                    Width =900
                    Height =255
                    FontSize =9
                    TabIndex =20
                    Name ="txtOnShelfDate"
                    Format ="Medium Time"

                    LayoutCachedLeft =10500
                    LayoutCachedTop =2940
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =3195
                End
            End
        End
        Begin FormFooter
            Height =690
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =11820
                    Height =630
                    FontSize =24
                    BackColor =65535
                    Name ="txtStatus"

                    LayoutCachedLeft =600
                    LayoutCachedTop =60
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =690
                End
            End
        End
    End
End
CodeBehindForm
' See "ScanBatchWithoutVoice.cls"
