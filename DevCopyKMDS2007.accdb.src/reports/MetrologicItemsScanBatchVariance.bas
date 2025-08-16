Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    ItemSuffix =44
    Left =1875
    Top =810
    RecSrcDt = Begin
        0xdeecad8c72f4e440
    End
    RecordSource ="qryMetrologicScanBatchVariance"
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
            ControlSource ="SalesDptDescr"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="CategoryDescr"
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
                    Left =-30
                    Width =10665
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Metrologic Scan Batch Variance"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =-30
                    LayoutCachedWidth =10635
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
                    ControlSource ="=\"Categories Between \" & Trim([Forms]![MetrologicOptimiser]![txtCategoryBegin]"
                        ") & \" And \" & [Forms]![MetrologicOptimiser]![txtCategoryBegin]"
                    Format ="Long Date"

                    LayoutCachedTop =840
                    LayoutCachedWidth =10621
                    LayoutCachedHeight =1140
                End
            End
        End
        Begin PageHeader
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =180
                    Width =1890
                    Height =225
                    Name ="Label5"
                    Caption ="Code"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =180
                    LayoutCachedWidth =2070
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextFontFamily =34
                    Left =2220
                    Width =2400
                    Height =225
                    Name ="Label7"
                    Caption ="Description"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =2220
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =4905
                    Width =1275
                    Height =225
                    Name ="Label9"
                    Caption ="Count"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =4905
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6420
                    Width =1275
                    Height =225
                    Name ="Label30"
                    Caption ="On Shelf"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7860
                    Width =1275
                    Height =225
                    Name ="Label32"
                    Caption ="Variance"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =7860
                    LayoutCachedWidth =9135
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =9300
                    Width =1275
                    Height =225
                    Name ="Label34"
                    Caption ="Variance Dollars"
                    FontName ="Arial Narrow"
                    LayoutCachedLeft =9300
                    LayoutCachedWidth =10575
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Top =60
                    Width =2880
                    Height =225
                    Name ="SalesDptDescr"
                    ControlSource ="SalesDptDescr"

                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader1"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =480
                    Top =60
                    Width =2880
                    Height =225
                    Name ="CategoryDescr"
                    ControlSource ="CategoryDescr"

                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =285
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
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =300
                    Width =1710
                    Height =225
                    Name ="Item.ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Item_ID"

                    LayoutCachedLeft =300
                    LayoutCachedWidth =2010
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2160
                    Width =2400
                    Height =225
                    TabIndex =1
                    Name ="Item.Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Item_Description"

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4680
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Count"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4680
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6240
                    Height =225
                    TabIndex =3
                    Name ="Text31"
                    ControlSource ="OnShelf"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6240
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7740
                    Height =225
                    TabIndex =4
                    Name ="Text33"
                    ControlSource ="Variance"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9240
                    Height =225
                    TabIndex =5
                    Name ="Text35"
                    ControlSource ="Variance Dollars"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9240
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =225
                    CurrencySymbol ="$"
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
            Height =390
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =120
                    Width =10500
                    Name ="Line19"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7740
                    Top =120
                    Height =270
                    FontSize =10
                    Name ="Text36"
                    ControlSource ="=Sum([Variance])"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =120
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9180
                    Top =120
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Text37"
                    ControlSource ="=Sum([Variance Dollars])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =9180
                    LayoutCachedTop =120
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =390
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4680
                    Top =120
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="Text40"
                    ControlSource ="=Sum([Count])"

                    LayoutCachedLeft =4680
                    LayoutCachedTop =120
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6240
                    Top =120
                    Height =270
                    FontSize =10
                    TabIndex =3
                    Name ="Text41"
                    ControlSource ="=Sum([OnShelf])"

                    LayoutCachedLeft =6240
                    LayoutCachedTop =120
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =960
                    Top =120
                    Height =270
                    FontSize =10
                    TabIndex =4
                    Name ="Text42"
                    ControlSource ="=Count([id])"

                    LayoutCachedLeft =960
                    LayoutCachedTop =120
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =390
                End
                Begin Label
                    Left =240
                    Top =120
                    Width =675
                    Height =240
                    Name ="Label43"
                    Caption ="* Total * "
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =915
                    LayoutCachedHeight =360
                End
            End
        End
    End
End
CodeBehindForm
' See "MetrologicItemsScanBatchVariance.cls"
