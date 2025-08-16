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
    Width =11460
    DatasheetFontHeight =11
    ItemSuffix =93
    Top =1485
    RecSrcDt = Begin
        0x05536bfd8f09e440
    End
    RecordSource ="qry66ScanBatchDetail"
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
            Height =510
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =-15
                    Width =11475
                    Height =510
                    FontSize =20
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="66 Scan Batch Detail by Item"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =-15
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =510
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
                    Left =2595
                    Top =144
                    Width =1020
                    Height =600
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label32"
                    Caption ="Physical Count"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =2595
                    LayoutCachedTop =144
                    LayoutCachedWidth =3615
                    LayoutCachedHeight =744
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =795
                    Width =11370
                    BorderColor =8388608
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =795
                    LayoutCachedWidth =11370
                    LayoutCachedHeight =795
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3675
                    Top =435
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

                    LayoutCachedLeft =3675
                    LayoutCachedTop =435
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =720
                End
                Begin Label
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =3609
                    Top =165
                    Width =1035
                    Height =585
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label64"
                    Caption ="OnShelf"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =3609
                    LayoutCachedTop =165
                    LayoutCachedWidth =4644
                    LayoutCachedHeight =750
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =75
                    Width =11370
                    BorderColor =8388608
                    Name ="Line67"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =75
                    LayoutCachedWidth =11370
                    LayoutCachedHeight =75
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =342
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Width =1446
                    Height =270
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedWidth =1446
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =1515
                    Width =2181
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1515
                    LayoutCachedWidth =3696
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =4035
                    Width =591
                    Height =270
                    ColumnWidth =2415
                    TabIndex =2
                    Name ="SumOfCountQuantity"
                    ControlSource ="CountQuantity"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4035
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4965
                    Width =606
                    Height =270
                    TabIndex =3
                    Name ="OnShelfAtTimeOfScanNOW"
                    ControlSource ="OnShelf"
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

                    LayoutCachedLeft =4965
                    LayoutCachedWidth =5571
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
                    Left =5832
                    Width =1179
                    Height =270
                    ColumnWidth =2145
                    TabIndex =4
                    Name ="FirstScan"
                    ControlSource ="SetupDate"
                    Format ="Long Time"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5832
                    LayoutCachedWidth =7011
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3750
                    Width =141
                    Height =270
                    TabIndex =5
                    Name ="OnShelfAtTimeOfScanNOWValue"
                    ControlSource ="Price"
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

                    LayoutCachedLeft =3750
                    LayoutCachedWidth =3891
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
                    IMESentenceMode =3
                    Left =7272
                    Width =2094
                    Height =270
                    TabIndex =6
                    Name ="Location"
                    ControlSource ="Location"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7272
                    LayoutCachedWidth =9366
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =0
                    IMESentenceMode =3
                    Left =9435
                    Height =270
                    TabIndex =7
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9435
                    LayoutCachedWidth =11136
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
                    Width =5625
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
                    LayoutCachedWidth =11325
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
                    Width =5625
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
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =11325
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
                    LayoutCachedWidth =11325
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Width =11325
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    BorderColor =0
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedWidth =11325
                    LayoutCachedHeight =255
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =855
            Name ="ReportFooter"
            Begin
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =75
                    Width =11370
                    BorderColor =8388608
                    Name ="Line69"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =75
                    LayoutCachedWidth =11370
                    LayoutCachedHeight =75
                End
                Begin Label
                    Left =5325
                    Top =210
                    Width =3750
                    Height =645
                    Name ="Label77"
                    Caption ="Note: On shelf = Onhand - Pending Transactions  (Reserved, Sold Today, Pending T"
                        "ransfers and Claims)"
                    LayoutCachedLeft =5325
                    LayoutCachedTop =210
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =855
                End
            End
        End
    End
End
CodeBehindForm
' See "66 Scan Batch Detail By Item.cls"
