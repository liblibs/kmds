Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =46
    Left =180
    Top =1350
    OrderBy ="ItemDescription"
    RecSrcDt = Begin
        0x8ff8b2133498e440
    End
    RecordSource ="qryMetrologicItemsNotScanned"
    Caption ="qryMetrologicItemsNotScanned"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SubCategory.Description"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =-60
                    Width =10725
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Metrologic Items With Onhands and Not Scanned "
                    FontName ="Times New Roman"
                    LayoutCachedLeft =-60
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =15
                    Top =540
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    Name ="Text84"
                    ControlSource ="=\"Sales Departments Between \" & Trim([Forms]![MetrologicOptimiser]![txtSalesDe"
                        "partmentBegin]) & \" And \" & [Forms]![MetrologicOptimiser]![txtSalesDepartmentB"
                        "egin]"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Top =840
                    Width =10621
                    Height =300
                    ColumnOrder =0
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"Categories Between \" & Trim([Forms]![MetrologicOptimiser]![cboNotScannedCate"
                        "goryIDBegin]) & \" And \" & [Forms]![MetrologicOptimiser]![cboNotScannedCategory"
                        "IDEnd]"
                    Format ="Long Date"

                    LayoutCachedTop =840
                    LayoutCachedWidth =10621
                    LayoutCachedHeight =1140
                End
            End
        End
        Begin PageHeader
            Height =225
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =1140
                    Width =1380
                    Height =225
                    Name ="Label1"
                    Caption ="Category"
                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Width =840
                    Height =225
                    Name ="Label3"
                    Caption ="Sales Dept"
                    LayoutCachedWidth =840
                    LayoutCachedHeight =225
                End
                Begin Label
                    Left =2880
                    Width =1050
                    Height =225
                    Name ="Label5"
                    Caption ="Code"
                    LayoutCachedLeft =2880
                    LayoutCachedWidth =3930
                    LayoutCachedHeight =225
                End
                Begin Label
                    Left =4080
                    Width =2400
                    Height =225
                    Name ="Label7"
                    Caption ="Description"
                    LayoutCachedLeft =4080
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =6765
                    Width =1275
                    Height =225
                    Name ="Label9"
                    Caption ="Onhand"
                    LayoutCachedLeft =6765
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =8160
                    Width =720
                    Height =225
                    Name ="Label30"
                    Caption ="Price"
                    LayoutCachedLeft =8160
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =9000
                    Width =720
                    Height =225
                    Name ="Label31"
                    Caption ="Extend"
                    LayoutCachedLeft =9000
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =225
            Name ="GroupHeader0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Width =480
                    Height =225
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"

                    LayoutCachedWidth =480
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =540
                    Width =2820
                    Height =225
                    TabIndex =1
                    Name ="Text36"
                    ControlSource ="SalesDepartment.Description"

                    LayoutCachedLeft =540
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =225
            BreakLevel =1
            Name ="GroupHeader1"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =840
                    Width =2460
                    Height =225
                    ColumnWidth =2160
                    Name ="SubCategory.Description"
                    ControlSource ="SubCategory.Description"
                    EventProcPrefix ="SubCategory_Description"

                    LayoutCachedLeft =840
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =225
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    OldBorderStyle =0
                    Left =60
                    Top =240
                    Width =9000
                    Name ="line5"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2520
                    Width =1050
                    Height =225
                    Name ="Item.ID"
                    ControlSource ="ItemID"
                    EventProcPrefix ="Item_ID"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3570
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3720
                    Width =2400
                    Height =225
                    TabIndex =1
                    Name ="Item.Description"
                    ControlSource ="ItemDescription"
                    EventProcPrefix ="Item_Description"

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    Format ="Standard"

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"

                    LayoutCachedLeft =8160
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9000
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text32"
                    ControlSource ="=[Onhand]*[Price]"
                    Format ="Standard"

                    LayoutCachedLeft =9000
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            BreakLevel =1
            Name ="GroupFooter0"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6240
                    Top =60
                    Width =1800
                    Height =255
                    FontSize =9
                    Name ="Text38"
                    ControlSource ="=Sum([Onhand])"
                    Format ="Standard"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =1080
                    Height =255
                    FontSize =9
                    TabIndex =1
                    Name ="Text39"
                    ControlSource ="=Sum([Onhand]*[Price])"
                    Format ="Standard"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =60
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =315
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    Left =3585
                    Top =60
                    Width =1770
                    Height =270
                    FontSize =9
                    Name ="Label40"
                    Caption ="Total        Count:"
                    LayoutCachedLeft =3585
                    LayoutCachedTop =60
                    LayoutCachedWidth =5355
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5400
                    Top =60
                    Width =540
                    Height =255
                    FontSize =9
                    TabIndex =2
                    Name ="Text41"
                    ControlSource ="=Count([Onhand])"
                    Format ="Standard"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =2460
                    Height =255
                    FontSize =9
                    TabIndex =3
                    Name ="Text42"
                    ControlSource ="SubCategory.Description"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =315
                End
                Begin Line
                    Top =420
                    Width =10500
                    Name ="Line43"
                    LayoutCachedTop =420
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =420
                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Top =300
                    Width =1920
                    Height =255
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5640
                    Top =300
                    Width =5040
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
                    Left =60
                    Width =4260
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =0
                    Left =5640
                    Width =5040
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Top =300
                    Width =10560
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =45
                    Width =10560
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =330
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =120
                    Width =10500
                    Name ="Line19"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6240
                    Top =60
                    Width =1800
                    Height =270
                    FontSize =10
                    Name ="Text33"
                    ControlSource ="=Sum([Onhand])"
                    Format ="Standard"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Text34"
                    ControlSource ="=Sum([Onhand]*[Price])"
                    Format ="Standard"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =60
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =330
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    TextFontFamily =34
                    Left =2880
                    Top =60
                    Width =2370
                    Height =270
                    FontSize =10
                    Name ="Label44"
                    Caption ="Total         Count:"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5295
                    Top =60
                    Width =825
                    Height =255
                    FontSize =10
                    TabIndex =2
                    Name ="Text45"
                    ControlSource ="=Count([Onhand])"
                    Format ="Standard"

                    LayoutCachedLeft =5295
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =315
                End
            End
        End
    End
End
CodeBehindForm
' See "MetrologicItemsNotScanned.cls"
