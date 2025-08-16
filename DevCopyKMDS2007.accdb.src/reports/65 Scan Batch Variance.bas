Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =20
    Width =15120
    DatasheetFontHeight =11
    ItemSuffix =92
    Top =810
    RecSrcDt = Begin
        0x6bed1a8377f6e340
    End
    RecordSource ="qryScanBatchVarianceCrossTab"
    Caption ="64 Kit Sales"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =204
            FontSize =9
            FontWeight =700
            BorderColor =5066944
            ForeColor =11830108
            FontName ="Calibri"
            GridlineColor =-2147483609
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
            BackColor =-2147483617
            BorderColor =-2147483617
            GridlineColor =-2147483609
        End
        Begin Line
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin Image
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackColor =-2147483617
            BorderColor =-2147483617
            GridlineColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =204
            Width =1701
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483609
            BorderLineStyle =0
        End
        Begin OptionButton
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            GridlineColor =-2147483609
        End
        Begin OptionGroup
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5066944
        End
        Begin BoundObjectFrame
            SizeMode =3
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =204
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            AsianLineBreak =1
            GridlineColor =-2147483609
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontCharSet =204
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =204
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            Width =4536
            Height =2835
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =204
            Width =283
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483609
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            BackStyle =0
            Width =5103
            Height =3402
            FontSize =9
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =1701
            Height =1701
            BorderColor =5066944
            GridlineColor =-2147483609
            LabelX =-1701
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Width =14970
                    Height =510
                    FontSize =20
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="65 Scan Batch Variance"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =510
                End
                Begin TextBox
                    DecimalPlaces =2
                    IMESentenceMode =3
                    Left =6732
                    Top =570
                    Width =1221
                    Height =270
                    ColumnOrder =0
                    Name ="Text78"
                    ControlSource ="=Sum(([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])*[Price])"
                    Format ="Standard"

                    LayoutCachedLeft =6732
                    LayoutCachedTop =570
                    LayoutCachedWidth =7953
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6297
                    Top =570
                    Width =711
                    Height =270
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=Sum([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])"

                    LayoutCachedLeft =6297
                    LayoutCachedTop =570
                    LayoutCachedWidth =7008
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6957
                    Top =855
                    Width =996
                    Height =270
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text80"
                    ControlSource ="=Sum(([OnShelfAtTimeOfScanNOW]*[Price]))"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000ccc8c200000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004f006e005300680065006c00660041007400540069006d0065004f006600 ,
                        0x5300630061006e004e004f0057005d003c003e005b004f006e00530068006500 ,
                        0x6c00660041007400540069006d0065004f0066005300630061006e005d000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =6957
                    LayoutCachedTop =855
                    LayoutCachedWidth =7953
                    LayoutCachedHeight =1125
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccc8c2002f0000005b00 ,
                        0x4f006e005300680065006c00660041007400540069006d0065004f0066005300 ,
                        0x630061006e004e004f0057005d003c003e005b004f006e005300680065006c00 ,
                        0x660041007400540069006d0065004f0066005300630061006e005d0000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin Label
                    Left =4107
                    Top =855
                    Width =2550
                    Height =285
                    Name ="Label81"
                    Caption ="Retail Value of Inventory on shelf"
                    LayoutCachedLeft =4107
                    LayoutCachedTop =855
                    LayoutCachedWidth =6657
                    LayoutCachedHeight =1140
                End
                Begin Label
                    Left =3885
                    Top =576
                    Width =2550
                    Height =285
                    Name ="Label82"
                    Caption ="Variance at Retail Value"
                    LayoutCachedLeft =3885
                    LayoutCachedTop =576
                    LayoutCachedWidth =6435
                    LayoutCachedHeight =861
                End
                Begin Label
                    Left =4107
                    Top =1140
                    Width =2550
                    Height =285
                    Name ="Label83"
                    Caption ="Variance/Retail value:"
                    LayoutCachedLeft =4107
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6657
                    LayoutCachedHeight =1425
                End
                Begin TextBox
                    DecimalPlaces =1
                    IMESentenceMode =3
                    Left =6762
                    Top =1155
                    Width =1221
                    Height =270
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Text84"
                    ControlSource ="=Sum(([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])*[Price])/Sum(([OnShelfAtTim"
                        "eOfScanNOW]*[Price]))"
                    Format ="Percent"

                    LayoutCachedLeft =6762
                    LayoutCachedTop =1155
                    LayoutCachedWidth =7983
                    LayoutCachedHeight =1425
                End
            End
        End
        Begin PageHeader
            Height =936
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Top =420
                    Width =825
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label116"
                    Caption ="Item ID"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedTop =420
                    LayoutCachedWidth =825
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =870
                    Top =420
                    Width =1155
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label117"
                    Caption ="Description"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =870
                    LayoutCachedTop =420
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =3315
                    Top =150
                    Width =930
                    Height =600
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label32"
                    Caption ="Physical Count"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =3315
                    LayoutCachedTop =150
                    LayoutCachedWidth =4245
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =13470
                    Top =420
                    Width =600
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label41"
                    Caption ="First"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =13470
                    LayoutCachedTop =420
                    LayoutCachedWidth =14070
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =14109
                    Top =420
                    Width =855
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label37"
                    Caption ="Last"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =14109
                    LayoutCachedTop =420
                    LayoutCachedWidth =14964
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =7059
                    Top =165
                    Width =1350
                    Height =585
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label40"
                    Caption ="OnShelf @ Time of Scan"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =7059
                    LayoutCachedTop =165
                    LayoutCachedWidth =8409
                    LayoutCachedHeight =750
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =795
                    Width =14970
                    BorderColor =8388608
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =795
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =795
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =5619
                    Top =420
                    Width =1305
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label60"
                    Caption ="Variance"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =5619
                    LayoutCachedTop =420
                    LayoutCachedWidth =6924
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =12885
                    Top =180
                    Width =2085
                    Height =570
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label61"
                    Caption ="Scans"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =12885
                    LayoutCachedTop =180
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =750
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4239
                    Top =465
                    Width =951
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8388608
                    Name ="txtONshelf"
                    ControlSource ="=Now()"
                    Format ="Medium Time"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =4239
                    LayoutCachedTop =465
                    LayoutCachedWidth =5190
                    LayoutCachedHeight =750
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =4239
                    Top =165
                    Width =945
                    Height =585
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label64"
                    Caption ="OnShelf"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =4239
                    LayoutCachedTop =165
                    LayoutCachedWidth =5184
                    LayoutCachedHeight =750
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =75
                    Width =14970
                    BorderColor =8388608
                    Name ="Line67"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =75
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =75
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =12885
                    Top =420
                    Width =630
                    Height =330
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label85"
                    Caption ="Count"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =12885
                    LayoutCachedTop =420
                    LayoutCachedWidth =13515
                    LayoutCachedHeight =750
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8490
                    Top =150
                    Width =1101
                    Height =570
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text54"
                    ControlSource ="=GetKMDSSettings(\"ScanBatchLoc1\",\"1. Store Shelf\")"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =8490
                    LayoutCachedTop =150
                    LayoutCachedWidth =9591
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9645
                    Top =150
                    Width =1101
                    Height =570
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text55"
                    ControlSource ="=GetKMDSSettings(\"ScanBatchLoc2\",\"2. Back Shelf\")"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =9645
                    LayoutCachedTop =150
                    LayoutCachedWidth =10746
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =10665
                    Top =150
                    Width =1026
                    Height =570
                    FontSize =11
                    FontWeight =700
                    TabIndex =3
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text56"
                    ControlSource ="=GetKMDSSettings(\"ScanBatchLoc3\",\"3. Overflow\")"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =10665
                    LayoutCachedTop =150
                    LayoutCachedWidth =11691
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =11805
                    Top =150
                    Width =891
                    Height =570
                    FontSize =11
                    FontWeight =700
                    TabIndex =4
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text57"
                    ControlSource ="=GetKMDSSettings(\"ScanBatchLoc4\",\"4. Display\")"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =11805
                    LayoutCachedTop =150
                    LayoutCachedWidth =12696
                    LayoutCachedHeight =720
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Width =1236
                    Height =270
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =1236
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1230
                    Width =2316
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1230
                    LayoutCachedWidth =3546
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =14190
                    Width =801
                    Height =270
                    ColumnWidth =2145
                    TabIndex =2
                    Name ="LastScan"
                    ControlSource ="LastScan"
                    Format ="Medium Time"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =14190
                    LayoutCachedWidth =14991
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =3600
                    Width =666
                    Height =270
                    ColumnWidth =2415
                    TabIndex =3
                    Name ="SumOfCountQuantity"
                    ControlSource ="SumOfCountQuantity"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3600
                    LayoutCachedWidth =4266
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =7059
                    Width =456
                    Height =270
                    ColumnWidth =2490
                    TabIndex =4
                    Name ="OnShelfAtTimeOfScan"
                    ControlSource ="OnShelfAtTimeOfScan"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7059
                    LayoutCachedWidth =7515
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4608
                    Width =561
                    Height =270
                    ColumnWidth =2985
                    TabIndex =5
                    Name ="OnShelfAtTimeOfScanNOW"
                    ControlSource ="OnShelfAtTimeOfScanNOW"
                    FontName ="Arial Narrow"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000ccc8c200000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004f006e005300680065006c00660041007400540069006d0065004f006600 ,
                        0x5300630061006e004e004f0057005d003c003e005b004f006e00530068006500 ,
                        0x6c00660041007400540069006d0065004f0066005300630061006e005d000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =4608
                    LayoutCachedWidth =5169
                    LayoutCachedHeight =270
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccc8c2002f0000005b00 ,
                        0x4f006e005300680065006c00660041007400540069006d0065004f0066005300 ,
                        0x630061006e004e004f0057005d003c003e005b004f006e005300680065006c00 ,
                        0x660041007400540069006d0065004f0066005300630061006e005d0000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13245
                    Width =876
                    Height =270
                    ColumnWidth =2145
                    TabIndex =6
                    Name ="FirstScan"
                    ControlSource ="FirstScan"
                    Format ="Medium Time"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =13245
                    LayoutCachedWidth =14121
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7554
                    Width =816
                    Height =270
                    TabIndex =7
                    Name ="LastOfOnShelfDate"
                    ControlSource ="LastOfOnShelfDate"
                    Format ="Medium Time"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7554
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5400
                    Width =606
                    Height =270
                    TabIndex =8
                    Name ="txtVariance"
                    ControlSource ="=[SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW]"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5400
                    LayoutCachedWidth =6006
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8709
                    Width =876
                    Height =270
                    TabIndex =9
                    Name ="Text50"
                    ControlSource ="Loc1"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8709
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9939
                    Width =816
                    Height =270
                    TabIndex =10
                    Name ="Text51"
                    ControlSource ="Loc2"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9939
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10809
                    Width =801
                    Height =270
                    TabIndex =11
                    Name ="Text52"
                    ControlSource ="Loc3"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =10809
                    LayoutCachedWidth =11610
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11949
                    Width =741
                    Height =270
                    TabIndex =12
                    Name ="Text53"
                    ControlSource ="Loc4"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =11949
                    LayoutCachedWidth =12690
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =6048
                    Width =891
                    Height =270
                    TabIndex =13
                    Name ="txtVarianceDollars"
                    ControlSource ="=([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])*[Price]"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6048
                    LayoutCachedWidth =6939
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4314
                    Width =576
                    Height =270
                    TabIndex =14
                    Name ="OnShelfAtTimeOfScanNOWValue"
                    ControlSource ="=([OnShelfAtTimeOfScanNOW]*[Price])"
                    FontName ="Arial Narrow"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000ccc8c200000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004f006e005300680065006c00660041007400540069006d0065004f006600 ,
                        0x5300630061006e004e004f0057005d003c003e005b004f006e00530068006500 ,
                        0x6c00660041007400540069006d0065004f0066005300630061006e005d000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =4314
                    LayoutCachedWidth =4890
                    LayoutCachedHeight =270
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccc8c2002f0000005b00 ,
                        0x4f006e005300680065006c00660041007400540069006d0065004f0066005300 ,
                        0x630061006e004e004f0057005d003c003e005b004f006e005300680065006c00 ,
                        0x660041007400540069006d0065004f0066005300630061006e005d0000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12810
                    Width =441
                    Height =270
                    TabIndex =15
                    Name ="txtScans"
                    ControlSource ="Scans"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =12810
                    LayoutCachedWidth =13251
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =3600
                    Height =255
                    FontWeight =700
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =9225
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =5700
                    LayoutCachedTop =300
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =7020
                    Height =255
                    FontWeight =700
                    TabIndex =2
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =60
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =9225
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =14925
                    Height =255
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedTop =300
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Width =14925
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    BorderColor =0
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedWidth =14925
                    LayoutCachedHeight =255
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1080
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    IMESentenceMode =3
                    Left =5760
                    Top =150
                    Width =1221
                    Height =270
                    Name ="txtTotalVariance"
                    ControlSource ="=Sum(([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])*[Price])"
                    Format ="Standard"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =150
                    LayoutCachedWidth =6981
                    LayoutCachedHeight =420
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =75
                    Width =14970
                    BorderColor =8388608
                    Name ="Line69"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =75
                    LayoutCachedWidth =14970
                    LayoutCachedHeight =75
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5250
                    Top =150
                    Width =711
                    Height =270
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])"

                    LayoutCachedLeft =5250
                    LayoutCachedTop =150
                    LayoutCachedWidth =5961
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =163
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5970
                    Top =435
                    Width =996
                    Height =270
                    TabIndex =2
                    Name ="txtTotalRetailValue"
                    ControlSource ="=Sum(([OnShelfAtTimeOfScanNOW]*[Price]))"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000ccc8c200000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004f006e005300680065006c00660041007400540069006d0065004f006600 ,
                        0x5300630061006e004e004f0057005d003c003e005b004f006e00530068006500 ,
                        0x6c00660041007400540069006d0065004f0066005300630061006e005d000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =5970
                    LayoutCachedTop =435
                    LayoutCachedWidth =6966
                    LayoutCachedHeight =705
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccc8c2002f0000005b00 ,
                        0x4f006e005300680065006c00660041007400540069006d0065004f0066005300 ,
                        0x630061006e004e004f0057005d003c003e005b004f006e005300680065006c00 ,
                        0x660041007400540069006d0065004f0066005300630061006e005d0000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin Label
                    Left =3168
                    Top =441
                    Width =2550
                    Height =285
                    Name ="Label73"
                    Caption ="Retail Value of Inventory on shelf"
                    LayoutCachedLeft =3168
                    LayoutCachedTop =441
                    LayoutCachedWidth =5718
                    LayoutCachedHeight =726
                End
                Begin Label
                    Left =3165
                    Top =150
                    Width =2550
                    Height =285
                    Name ="Label74"
                    Caption ="Variance at Retail Value"
                    LayoutCachedLeft =3165
                    LayoutCachedTop =150
                    LayoutCachedWidth =5715
                    LayoutCachedHeight =435
                End
                Begin Label
                    Left =3168
                    Top =726
                    Width =2550
                    Height =285
                    Name ="Label75"
                    Caption ="Variance/Retail value:"
                    LayoutCachedLeft =3168
                    LayoutCachedTop =726
                    LayoutCachedWidth =5718
                    LayoutCachedHeight =1011
                End
                Begin TextBox
                    DecimalPlaces =1
                    IMESentenceMode =3
                    Left =5685
                    Top =720
                    Width =1221
                    Height =270
                    TabIndex =3
                    Name ="txtTotalVariancePercent"
                    ControlSource ="=Sum(([SumOfCountQuantity]-[OnShelfAtTimeOfScanNOW])*[Price])/Sum(([OnShelfAtTim"
                        "eOfScanNOW]*[Price]))"
                    Format ="Percent"

                    LayoutCachedLeft =5685
                    LayoutCachedTop =720
                    LayoutCachedWidth =6906
                    LayoutCachedHeight =990
                End
                Begin Label
                    Left =8568
                    Top =216
                    Width =3750
                    Height =645
                    Name ="Label77"
                    Caption ="Note: On shelf = Onhand - Pending Transactions  (Reserved, Sold Today, Pending T"
                        "ransfers and Claims)"
                    LayoutCachedLeft =8568
                    LayoutCachedTop =216
                    LayoutCachedWidth =12318
                    LayoutCachedHeight =861
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =12525
                    Top =150
                    Width =786
                    Height =270
                    TabIndex =4
                    Name ="Text88"
                    ControlSource ="=Sum([scans])"

                    LayoutCachedLeft =12525
                    LayoutCachedTop =150
                    LayoutCachedWidth =13311
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =150
                    Top =210
                    Width =366
                    Height =270
                    TabIndex =5
                    BackColor =12765388
                    Name ="txtHilite"
                    ControlSource ="=\"*\""

                    LayoutCachedLeft =150
                    LayoutCachedTop =210
                    LayoutCachedWidth =516
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            Left =570
                            Top =210
                            Width =2100
                            Height =855
                            Name ="Label91"
                            Caption ="NOTE: Gray \"On Shelf\" totals have changed.  Please very physical count."
                            LayoutCachedLeft =570
                            LayoutCachedTop =210
                            LayoutCachedWidth =2670
                            LayoutCachedHeight =1065
                        End
                    End
                End
            End
        End
    End
End
