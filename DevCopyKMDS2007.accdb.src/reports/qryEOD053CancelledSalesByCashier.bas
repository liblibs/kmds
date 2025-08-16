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
    ItemSuffix =85
    Top =615
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xa926d6663a0be340
    End
    RecordSource ="qryInvoLine5CancelledSalesDetail"
    Caption ="EOD Sales Returns By Cashier By Item"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="InvoHeadID"
        End
        Begin FormHeader
            Height =600
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-60
                    Width =9285
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Cancelled Sales By Cashier By Invoice"
                End
            End
        End
        Begin PageHeader
            Height =480
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =120
                    Top =60
                    Width =780
                    Height =300
                    Name ="Name_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1080
                    Top =60
                    Width =495
                    Height =300
                    Name ="TotalTransactions_Label"
                    Caption ="Date"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =10560
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10560
                    Name ="Line42"
                End
                Begin Label
                    Left =4560
                    Top =60
                    Width =1515
                    Height =300
                    Name ="Label46"
                    Caption ="Cancelled Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8010
                    Top =60
                    Width =1005
                    Height =300
                    Name ="Label57"
                    Caption =" Quantity "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9810
                    Top =60
                    Width =810
                    Height =300
                    Name ="Label58"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2760
                    Top =60
                    Width =1605
                    Height =300
                    Name ="Label74"
                    Caption ="Invoice "
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =2025
                    Top =60
                    Width =615
                    Height =300
                    Name ="Label84"
                    Caption ="Reg #"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2025
                    LayoutCachedTop =60
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =360
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    FontWeight =700
                    Name ="Text75"
                    ControlSource ="CashierName"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2760
                    Width =1740
                    Height =225
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2760
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =45
                    Width =2055
                    ColumnWidth =2100
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2100
                    Width =540
                    Height =225
                    TabIndex =2
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2100
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =225
                End
            End
        End
        Begin Section
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3960
                    Height =225
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =3960
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5460
                    Width =2820
                    Height =225
                    ColumnWidth =3150
                    TabIndex =1
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
                    TabIndex =2
                    Name ="SumOfQuantity"
                    ControlSource ="Quantity"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9120
                    Height =225
                    TabIndex =3
                    Name ="SumOfPrice"
                    ControlSource ="Price"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2520
                    Height =225
                    TabIndex =4
                    Name ="InputCode"
                    ControlSource ="InputCode"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    FontWeight =700
                    Name ="Text77"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =1
                    Name ="Text78"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    BorderWidth =1
                    Left =8280
                    Top =60
                    Width =2280
                    Name ="Line79"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =120
                    FontWeight =700
                    Name ="CashierName"
                    ControlSource ="CashierName"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    BorderWidth =1
                    Top =420
                    Width =10560
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =1
                    Left =6900
                    Top =60
                    Width =3660
                    Name ="Line73"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    FontWeight =700
                    TabIndex =1
                    Name ="Text80"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =2
                    Name ="Text81"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Arial Narrow"

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
            Height =480
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
                    FontWeight =700
                    Name ="Text82"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =1
                    Name ="Text83"
                    ControlSource ="=Sum([Quantity])"
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD053CancelledSalesByCashier.cls"
