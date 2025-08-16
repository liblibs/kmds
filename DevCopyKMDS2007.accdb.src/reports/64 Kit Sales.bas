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
    GridY =10
    Width =10485
    DatasheetFontHeight =11
    ItemSuffix =35
    Left =855
    Top =690
    RecSrcDt = Begin
        0x0ed89abbcee7e340
    End
    RecordSource ="qryKitsWithTotalsAndItems"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="KitItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =582
            Name ="ReportHeader"
            AutoHeight =1
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Width =10485
                    Height =510
                    FontSize =20
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Kits on Sale"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =510
                End
            End
        End
        Begin PageHeader
            Height =375
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Width =10410
                    BorderColor =8388608
                    Name ="Line55"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedWidth =10410
                End
                Begin Line
                    OldBorderStyle =1
                    BorderWidth =2
                    Top =360
                    Width =10410
                    BorderColor =8388608
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    GridlineColor =0
                    LayoutCachedTop =360
                    LayoutCachedWidth =10410
                    LayoutCachedHeight =360
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =390
                    Top =56
                    Width =1200
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label116"
                    Caption ="Kit Item ID"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =390
                    LayoutCachedTop =56
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =356
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =1700
                    Top =56
                    Width =1485
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label117"
                    Caption ="Description"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =1700
                    LayoutCachedTop =56
                    LayoutCachedWidth =3185
                    LayoutCachedHeight =356
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =4815
                    Top =60
                    Width =1275
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label118"
                    Caption ="Selling Price"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =4815
                    LayoutCachedTop =60
                    LayoutCachedWidth =6090
                    LayoutCachedHeight =360
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =8235
                    Top =56
                    Width =1005
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label119"
                    Caption ="Discount"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =8235
                    LayoutCachedTop =56
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =356
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =18
                    Left =6519
                    Top =56
                    Width =1275
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label32"
                    Caption ="Retail Value"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =6519
                    LayoutCachedTop =56
                    LayoutCachedWidth =7794
                    LayoutCachedHeight =356
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =9297
                    Top =56
                    Width =1125
                    Height =300
                    FontSize =11
                    BorderColor =0
                    ForeColor =8388608
                    Name ="Label33"
                    Caption ="Discount %"
                    FontName ="Times New Roman"
                    GridlineColor =0
                    LayoutCachedLeft =9297
                    LayoutCachedTop =56
                    LayoutCachedWidth =10422
                    LayoutCachedHeight =356
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =570
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Top =225
                    Width =1590
                    Height =345
                    FontSize =12
                    FontWeight =700
                    ForeColor =11830108
                    Name ="qryKitsOnFile.ItemID"
                    ControlSource ="KitItemID"
                    EventProcPrefix ="qryKitsOnFile_ItemID"
                    GridlineColor =0

                    LayoutCachedTop =225
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1700
                    Top =225
                    Width =3075
                    Height =345
                    FontSize =12
                    TabIndex =1
                    Name ="qryKitsOnFile.Description"
                    ControlSource ="KitDescription"
                    EventProcPrefix ="qryKitsOnFile_Description"

                    LayoutCachedLeft =1700
                    LayoutCachedTop =225
                    LayoutCachedWidth =4775
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4818
                    Top =225
                    Width =1500
                    Height =345
                    FontSize =12
                    TabIndex =2
                    Name ="qryKitsOnFile.Price"
                    ControlSource ="KitPrice"
                    EventProcPrefix ="qryKitsOnFile_Price"

                    LayoutCachedLeft =4818
                    LayoutCachedTop =225
                    LayoutCachedWidth =6318
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6483
                    Top =225
                    Width =1311
                    Height =345
                    FontSize =12
                    TabIndex =3
                    Name ="RetailPrice"
                    ControlSource ="RetailPrice"

                    LayoutCachedLeft =6483
                    LayoutCachedTop =225
                    LayoutCachedWidth =7794
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8049
                    Top =225
                    Width =1191
                    Height =345
                    FontSize =12
                    TabIndex =4
                    Name ="Discount"
                    ControlSource ="KitDiscount"

                    LayoutCachedLeft =8049
                    LayoutCachedTop =225
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9297
                    Top =225
                    Width =966
                    Height =345
                    ColumnWidth =2205
                    FontSize =12
                    TabIndex =5
                    Name ="KitDiscPerct"
                    ControlSource ="KitDiscPerct"
                    Format ="Percent"
                    ConditionalFormat = Begin
                        0x0100000072000000020000000000000004000000000000000500000001000000 ,
                        0xed1c2400ffffff00000000000500000006000000080000000100000000000000 ,
                        0xfff2000000000000000000000000000000000000000000000000000000000000 ,
                        0x30002e003100350000000000300000000000
                    End

                    LayoutCachedLeft =9297
                    LayoutCachedTop =225
                    LayoutCachedWidth =10263
                    LayoutCachedHeight =570
                    ConditionalFormat14 = Begin
                        0x010002000000000000000400000001000000ed1c2400ffffff00040000003000 ,
                        0x2e00310035000000000000000000000000000000000000000000000000000005 ,
                        0x0000000100000000000000fff200000100000030000000000000000000000000 ,
                        0x00000000000000000000
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =1755
                    Top =120
                    Width =8509
                    Height =18
                    Name ="Line34"
                    LayoutCachedLeft =1755
                    LayoutCachedTop =120
                    LayoutCachedWidth =10264
                    LayoutCachedHeight =138
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
                    IMESentenceMode =3
                    Left =3622
                    Width =2610
                    Height =270
                    Name ="PartDescription"
                    ControlSource ="PartDescription"

                    LayoutCachedLeft =3622
                    LayoutCachedWidth =6232
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2324
                    Width =1245
                    Height =270
                    TabIndex =1
                    Name ="PartID"
                    ControlSource ="PartID"

                    LayoutCachedLeft =2324
                    LayoutCachedWidth =3569
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6405
                    Width =1365
                    Height =270
                    TabIndex =2
                    Name ="PartPrice"
                    ControlSource ="PartPrice"

                    LayoutCachedLeft =6405
                    LayoutCachedWidth =7770
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
                    Width =1920
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
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =4740
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
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =5340
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
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =4740
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
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Top =300
                    Width =10380
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
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =555
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =180
                    Width =10260
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    BorderColor =0
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"
                    GridlineColor =0

                    LayoutCachedLeft =180
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =255
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
