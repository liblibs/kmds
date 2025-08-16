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
    ItemSuffix =145
    Left =1560
    Top =135
    OrderBy ="ID"
    RecSrcDt = Begin
        0x2616e4479369e440
    End
    RecordSource ="qryItemSalesByCashierReport"
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
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="SalesDptID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1800
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
                    Caption ="Item Sales by Cashier"
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
            Height =600
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =180
                    Top =120
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =1335
                    LayoutCachedHeight =420
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
                    Left =6240
                    Top =120
                    Width =1740
                    Height =300
                    Name ="Label86"
                    Caption ="Register - Invoice"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6240
                    LayoutCachedTop =120
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9390
                    Top =120
                    Width =1230
                    Height =300
                    Name ="Label87"
                    Caption ="Dollars Sold"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9390
                    LayoutCachedTop =120
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10560
                    Name ="Line57"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextAlign =3
                    Left =8175
                    Top =120
                    Width =945
                    Height =300
                    Name ="Label113"
                    Caption ="Quantity "
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8175
                    LayoutCachedTop =120
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =420
                End
                Begin Label
                    Left =4620
                    Top =120
                    Width =1005
                    Height =300
                    Name ="Label118"
                    Caption ="Sold on"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4620
                    LayoutCachedTop =120
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =420
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =480
                    Height =225
                    Name ="CashierID"
                    ControlSource ="CashierID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =720
                    Top =60
                    Height =225
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
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
                Begin TextBox
                    TextAlign =1
                    Left =1440
                    Width =2394
                    ColumnWidth =1185
                    Name ="GroupDesc"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =3834
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    Left =45
                    Width =1335
                    TabIndex =1
                    Name ="Group"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Width =900
                    ColumnWidth =1770
                    TabIndex =2
                    Name ="SumOfPrice"
                    ControlSource ="Price"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9720
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Width =780
                    TabIndex =3
                    Name ="SumOfQuantity"
                    ControlSource ="Quantity"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8340
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4140
                    Width =2040
                    TabIndex =4
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="General Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =240
                End
                Begin Line
                    Visible = NotDefault
                    Left =3600
                    Top =240
                    Width =7080
                    Name ="Line5"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Width =480
                    TabIndex =5
                    Name ="Text143"
                    ControlSource ="Register"
                    Format ="Fixed"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6480
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7080
                    Width =1245
                    TabIndex =6
                    Name ="Text144"
                    ControlSource ="Invoice"
                    Format ="Fixed"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =8325
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter0"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Width =900
                    Height =330
                    Name ="Text133"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"

                    LayoutCachedLeft =9720
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Width =780
                    Height =330
                    TabIndex =1
                    Name ="Text135"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"

                    LayoutCachedLeft =8340
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4500
                    Width =720
                    TabIndex =2
                    Name ="SaleDpt"
                    ControlSource ="SalesDptID"

                    LayoutCachedLeft =4500
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5280
                    Width =3060
                    TabIndex =3
                    Name ="SaleDptDesc"
                    ControlSource ="SalesDptDesc"

                    LayoutCachedLeft =5280
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Width =900
                    Height =330
                    Name ="Text138"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"

                    LayoutCachedLeft =9720
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            Left =3000
                            Width =1095
                            Height =330
                            Name ="Label139"
                            Caption =" Total For:"
                            LayoutCachedLeft =3000
                            LayoutCachedWidth =4095
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Width =780
                    Height =330
                    TabIndex =1
                    Name ="Text140"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"

                    LayoutCachedLeft =8340
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4440
                    Width =720
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text141"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4440
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5220
                    Width =3060
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text142"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5220
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =300
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
            Height =840
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9720
                    Top =120
                    Width =900
                    Name ="Text122"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"

                    LayoutCachedLeft =9720
                    LayoutCachedTop =120
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =360
                    Begin
                        Begin Label
                            Left =4740
                            Top =120
                            Width =2040
                            Height =300
                            Name ="Label124"
                            Caption ="Report Total"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =780
                    TabIndex =1
                    Name ="Text123"
                    ControlSource ="=Sum([Quantity])"
                    Format ="Standard"

                    LayoutCachedLeft =8340
                    LayoutCachedTop =120
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =360
                End
            End
        End
    End
End
CodeBehindForm
' See "70 Item Sales By Cashier.cls"
