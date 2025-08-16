Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10785
    DatasheetFontHeight =10
    ItemSuffix =165
    Left =1935
    Top =840
    OrderBy ="SumOfPrice DESC"
    RecSrcDt = Begin
        0x73ba2ceeafc2e540
    End
    RecordSource ="qryItemSalesByDepartmentByVendorReport"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="VendorName"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SalesDptID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1980
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =420
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-60
                    Top =60
                    Width =10845
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Item Sales By Vendor By Department"
                    LayoutCachedLeft =-60
                    LayoutCachedTop =60
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10680
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1200
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Sales based on Sales Between \" & Format(Forms!PrintInventoryReport!txtDateRa"
                        "ngeFrom,\"mm/dd/yyyy\") & \" And \" & Format(Forms!PrintInventoryReport!txtDateR"
                        "angeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1500
                    Width =10621
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =555
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =180
                    Top =180
                    Width =510
                    Height =330
                    Name ="Department_Label"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =180
                    LayoutCachedTop =180
                    LayoutCachedWidth =690
                    LayoutCachedHeight =510
                End
                Begin Label
                    TextAlign =3
                    Left =4170
                    Top =180
                    Width =570
                    Height =330
                    Name ="TotalP_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4170
                    LayoutCachedTop =180
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =510
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =120
                    Width =10620
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6480
                    Top =180
                    Width =825
                    Height =330
                    Name ="Label87"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6480
                    LayoutCachedTop =180
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =510
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =540
                    Width =10560
                    Name ="Line57"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =540
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =540
                End
                Begin Label
                    TextAlign =3
                    Left =5220
                    Top =180
                    Width =945
                    Height =330
                    Name ="Label113"
                    Caption ="Sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5220
                    LayoutCachedTop =180
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =510
                End
                Begin Label
                    TextAlign =3
                    Left =7920
                    Top =180
                    Width =825
                    Height =330
                    Name ="Label148"
                    Caption ="Cost"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7920
                    LayoutCachedTop =180
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =510
                End
                Begin Label
                    TextAlign =3
                    Left =10200
                    Top =180
                    Width =540
                    Height =330
                    Name ="Label149"
                    Caption ="G.P."
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =10200
                    LayoutCachedTop =180
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =510
                End
                Begin Label
                    TextAlign =3
                    Left =9060
                    Top =180
                    Width =825
                    Height =330
                    Name ="Label160"
                    Caption ="Profit"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =180
                    LayoutCachedWidth =9885
                    LayoutCachedHeight =510
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =480
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =2940
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="VendorName"
                    ControlSource ="VendorName"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =0
                            Top =60
                            Width =840
                            Height =300
                            Name ="Label142"
                            Caption ="Vendor:"
                            LayoutCachedTop =60
                            LayoutCachedWidth =840
                            LayoutCachedHeight =360
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Width =720
                    Name ="SaleDpt"
                    ControlSource ="SalesDptID"

                    LayoutCachedWidth =720
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =780
                    Width =3060
                    TabIndex =1
                    Name ="SaleDptDesc"
                    ControlSource ="SalesDptDesc"

                    LayoutCachedLeft =780
                    LayoutCachedWidth =3840
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
                Begin TextBox
                    TextAlign =1
                    Left =1380
                    Width =2394
                    ColumnWidth =1185
                    FontWeight =700
                    Name ="GroupDesc"
                    ControlSource ="Description"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =3774
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    Left =45
                    Width =1275
                    FontWeight =700
                    TabIndex =1
                    Name ="Group"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4020
                    Width =720
                    Height =225
                    ColumnWidth =885
                    TabIndex =2
                    Name ="Price"
                    ControlSource ="Price"

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =12
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6585
                    Width =960
                    ColumnWidth =1770
                    TabIndex =3
                    Name ="SumOfPrice"
                    ControlSource ="SumOfPrice"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6585
                    LayoutCachedWidth =7545
                    LayoutCachedHeight =240
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5445
                    Width =780
                    TabIndex =4
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"
                    Format ="Standard"

                    LayoutCachedLeft =5445
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =240
                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10680
                    Name ="Line5"
                    LayoutCachedTop =240
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10020
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Profit"
                    ControlSource ="Profit"
                    Format ="Percent"

                    LayoutCachedLeft =10020
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7725
                    Width =1020
                    TabIndex =6
                    Name ="SumOfCost"
                    ControlSource ="SumOfCost"

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8880
                    Width =1020
                    TabIndex =7
                    Name ="Text161"
                    ControlSource ="=[SumOfPrice]-[SumOfCost]"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8880
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =240
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =630
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6585
                    Width =960
                    FontWeight =700
                    Name ="Text138"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6585
                    LayoutCachedWidth =7545
                    LayoutCachedHeight =240
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5445
                    Width =780
                    TabIndex =1
                    Name ="Text139"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                    LayoutCachedLeft =5445
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =2520
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text140"
                    ControlSource ="SalesDptDesc"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =4500
                    Top =60
                    Width =600
                    Height =300
                    Name ="Label141"
                    Caption ="Total"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =60
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Top =60
                    Width =1800
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text143"
                    ControlSource ="VendorName"
                    FontName ="Times New Roman"

                    LayoutCachedTop =60
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7725
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =4
                    Name ="Text150"
                    ControlSource ="=Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7725
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =255
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Width =675
                    Height =225
                    FontWeight =700
                    TabIndex =5
                    Name ="Text154"
                    ControlSource ="=((Sum([SumOfPrice])-Sum([SumOfCost]))/Sum([SumOfPrice]))"
                    Format ="Percent"

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8880
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =6
                    Name ="Text162"
                    ControlSource ="=Sum([SumOfPrice])-Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8880
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =255
                    CurrencySymbol ="$"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =690
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6585
                    Top =60
                    Width =960
                    FontWeight =700
                    Name ="Text131"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6585
                    LayoutCachedTop =60
                    LayoutCachedWidth =7545
                    LayoutCachedHeight =300
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5445
                    Top =60
                    Width =780
                    TabIndex =1
                    Name ="Text133"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                    LayoutCachedLeft =5445
                    LayoutCachedTop =60
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =780
                    Top =60
                    Width =2760
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text135"
                    ControlSource ="VendorName"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =3600
                    Top =60
                    Width =1260
                    Height =300
                    Name ="Label137"
                    Caption ="Total"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7725
                    Top =60
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="Text152"
                    ControlSource ="=Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7725
                    LayoutCachedTop =60
                    LayoutCachedWidth =8745
                    LayoutCachedHeight =315
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =60
                    Width =675
                    Height =225
                    FontWeight =700
                    TabIndex =4
                    Name ="Text158"
                    ControlSource ="=((Sum([SumOfPrice])-Sum([SumOfCost]))/Sum([SumOfPrice]))"
                    Format ="Percent"

                    LayoutCachedLeft =9960
                    LayoutCachedTop =60
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8880
                    Top =60
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    Name ="Text163"
                    ControlSource ="=Sum([SumOfPrice])-Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =60
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =315
                    CurrencySymbol ="$"
                End
            End
        End
        Begin PageFooter
            Height =585
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
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
                    Left =5700
                    Top =300
                    Width =4965
                    Height =285
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
                    Left =60
                    Width =4380
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
                    Left =5700
                    Width =4965
                    Height =285
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
                    Top =300
                    Width =10680
                    Height =255
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
                    Width =10620
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
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6180
                    Top =120
                    Width =1380
                    FontWeight =700
                    Name ="Text122"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =360
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            Left =1320
                            Top =120
                            Width =2040
                            Height =300
                            Name ="Label124"
                            Caption ="Report Total"
                            LayoutCachedLeft =1320
                            LayoutCachedTop =120
                            LayoutCachedWidth =3360
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5340
                    Top =120
                    Width =960
                    TabIndex =1
                    Name ="Text123"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"

                    LayoutCachedLeft =5340
                    LayoutCachedTop =120
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7740
                    Top =120
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =2
                    Name ="Text153"
                    ControlSource ="=Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =7740
                    LayoutCachedTop =120
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =375
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9960
                    Top =120
                    Width =675
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="Text159"
                    ControlSource ="=((Sum([SumOfPrice])-Sum([SumOfCost]))/Sum([SumOfPrice]))"
                    Format ="Percent"

                    LayoutCachedLeft =9960
                    LayoutCachedTop =120
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8880
                    Top =120
                    Width =1020
                    Height =255
                    FontWeight =700
                    TabIndex =4
                    Name ="Text164"
                    ControlSource ="=Sum([SumOfPrice])-Sum([SumOfCost])"
                    Format ="$#,##0.00;($#,##0.00)"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =120
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =375
                    CurrencySymbol ="$"
                End
            End
        End
    End
End
CodeBehindForm
' See "77 Item Sales By Department By Vendor.cls"
