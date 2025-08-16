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
    ItemSuffix =75
    Left =525
    Top =1230
    RecSrcDt = Begin
        0x4bfa8e546bcae240
    End
    RecordSource ="qryEOSalesReturnsByCashierByItem"
    Caption ="Sales Returns By Cashier By Item"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
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
            GroupFooter = NotDefault
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            ControlSource ="SetupDate"
        End
        Begin FormHeader
            Height =1080
            Name ="ReportHeader"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =540
                    Width =10560
                    Name ="Line41"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-45
                    Width =9255
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption =" Sales Returns By Cashier By Item"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Top =600
                    Width =10621
                    Height =480
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin PageHeader
            Height =675
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =7560
                    Top =300
                    Width =780
                    Height =300
                    Name ="Name_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1207
                    Top =300
                    Width =1230
                    Height =300
                    Name ="TotalTransactions_Label"
                    Caption ="Invoice Date"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =660
                    Width =10560
                    Name ="Line42"
                End
                Begin Label
                    Left =4500
                    Top =300
                    Width =1395
                    Height =300
                    Name ="Label46"
                    Caption ="Returned Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8760
                    Width =780
                    Height =600
                    Name ="Label57"
                    Caption ="Return Qty "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9750
                    Width =810
                    Height =600
                    Name ="Label58"
                    Caption ="Return Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2520
                    Top =300
                    Width =1575
                    Height =300
                    Name ="Label74"
                    Caption ="Invoice Number"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2760
                    Width =240
                    Height =225
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3060
                    Width =1020
                    Height =225
                    TabIndex =1
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3060
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4140
                    Width =1260
                    Height =225
                    TabIndex =2
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5460
                    Width =2820
                    Height =225
                    ColumnWidth =3150
                    TabIndex =3
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Width =660
                    Height =225
                    TabIndex =4
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Height =225
                    TabIndex =5
                    Name ="SumOfPrice"
                    ControlSource ="SumOfPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =615
                    Width =2085
                    ColumnWidth =2100
                    TabIndex =6
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =615
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =435
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Top =120
                    Width =300
                    FontWeight =700
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =120
                    FontWeight =700
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    FontWeight =700
                    TabIndex =2
                    Name ="Text67"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =3
                    Name ="Text68"
                    ControlSource ="=Sum([SumOfQuantity])"

                End
                Begin Line
                    BorderWidth =1
                    Top =420
                    Width =10560
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =1
                    Left =6480
                    Top =60
                    Width =4080
                    Name ="Line73"
                End
            End
        End
        Begin PageFooter
            Height =600
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Top =300
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5640
                    Top =300
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Width =5040
                    Height =300
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
                    Left =5640
                    Width =5040
                    Height =300
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
                    Height =300
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
                    Width =10680
                    Height =255
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
            Height =420
            Name ="ReportFooter"
            Begin
                Begin Line
                    Top =60
                    Width =10560
                    Name ="Line45"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    Height =300
                    FontWeight =700
                    Name ="Text71"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text72"
                    ControlSource ="=Sum([SumOfQuantity])"

                End
            End
        End
    End
End
CodeBehindForm
' See "30 SalesReturnsByCashier.cls"
