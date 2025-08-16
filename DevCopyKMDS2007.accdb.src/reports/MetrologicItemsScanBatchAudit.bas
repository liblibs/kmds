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
    ItemSuffix =47
    Left =2235
    Top =780
    RecSrcDt = Begin
        0xc6cde52db397e440
    End
    RecordSource ="qryMetroScanBatchDetail"
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
                    Left =-45
                    Width =10695
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Metrologic Scan Batch Log Report"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =-45
                    LayoutCachedWidth =10650
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
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Width =840
                    Height =225
                    Name ="Label5"
                    Caption ="Code"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =225
                End
                Begin Label
                    Left =960
                    Width =2400
                    Height =225
                    Name ="Label7"
                    Caption ="Description"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =3525
                    Width =1275
                    Height =225
                    Name ="Label9"
                    Caption ="Scan Code"
                    LayoutCachedLeft =3525
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =5100
                    Width =720
                    Height =225
                    Name ="Label30"
                    Caption ="Quantity"
                    LayoutCachedLeft =5100
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =6060
                    Width =660
                    Height =225
                    Name ="Label32"
                    Caption ="Price"
                    LayoutCachedLeft =6060
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =8340
                    Width =1275
                    Height =225
                    Name ="Label34"
                    Caption ="Date"
                    LayoutCachedLeft =8340
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =34
                    Left =7080
                    Width =900
                    Height =225
                    Name ="Label40"
                    Caption ="Extend"
                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =225
                End
                Begin Line
                    Top =240
                    Width =10500
                    Name ="Line46"
                    LayoutCachedTop =240
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =240
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
                    Width =840
                    Height =225
                    Name ="Item.ID"
                    ControlSource ="ID"
                    EventProcPrefix ="Item_ID"

                    LayoutCachedWidth =840
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Width =2400
                    Height =225
                    TabIndex =1
                    Name ="Item.Description"
                    ControlSource ="Description"
                    EventProcPrefix ="Item_Description"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3360
                    Height =225
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="ScanCode"

                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4935
                    Width =885
                    Height =225
                    TabIndex =3
                    Name ="Text31"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =4935
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5940
                    Width =840
                    Height =225
                    TabIndex =4
                    Name ="Text33"
                    ControlSource ="Price"
                    Format ="Fixed"

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8040
                    Width =1200
                    Height =225
                    TabIndex =5
                    Name ="Text35"
                    ControlSource ="ScanDate"
                    Format ="Short Date"

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9360
                    Width =1260
                    Height =225
                    TabIndex =6
                    Name ="Text39"
                    ControlSource ="ScanTime"
                    Format ="Long Time"

                    LayoutCachedLeft =9360
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6900
                    Width =1080
                    Height =225
                    TabIndex =7
                    Name ="Text41"
                    ControlSource ="=[Price]*[Quantity]"
                    Format ="Standard"

                    LayoutCachedLeft =6900
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =225
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
            Height =420
            Name ="ReportFooter"
            Begin
                Begin Line
                    Left =120
                    Top =60
                    Width =10500
                    Name ="Line19"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =60
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4935
                    Top =60
                    Width =885
                    Height =270
                    FontSize =10
                    Name ="Text42"
                    ControlSource ="=Sum([Quantity])"

                    LayoutCachedLeft =4935
                    LayoutCachedTop =60
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6900
                    Top =60
                    Width =1080
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Text43"
                    ControlSource ="=Sum([Price]*[Quantity])"
                    Format ="Standard"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =60
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =330
                End
            End
        End
    End
End
CodeBehindForm
' See "MetrologicItemsScanBatchAudit.cls"
