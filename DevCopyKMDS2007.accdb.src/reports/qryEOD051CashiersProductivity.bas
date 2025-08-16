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
    ItemSuffix =59
    Top =555
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xa5777a14e9c8e240
    End
    RecordSource ="qryEOCashiersProductivityReport"
    Caption ="qryInvoLine5SalesByCashierSummary"
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
        Begin FormHeader
            Height =1320
            Name ="ReportHeader"
            Begin
                Begin Label
                    Left =780
                    Top =900
                    Width =780
                    Height =300
                    Name ="Name_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2550
                    Top =900
                    Width =1050
                    Height =300
                    Name ="TotalTransactions_Label"
                    Caption ="Customers"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4230
                    Top =900
                    Width =720
                    Height =300
                    Name ="TotalSales_Label"
                    Caption ="Sales *"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5070
                    Top =900
                    Width =870
                    Height =300
                    Name ="AvgSale_Label"
                    Caption ="Avg Sale"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6240
                    Top =660
                    Width =570
                    Height =555
                    Name ="TotalNoSales_Label"
                    Caption ="No Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6847
                    Top =915
                    Width =975
                    Height =300
                    Name ="TotalCancelledSales_Label"
                    Caption ="Cancelled"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =780
                    Top =540
                    Width =9780
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =2
                    Left =780
                    Top =1260
                    Width =9780
                    Name ="Line42"
                End
                Begin Label
                    TextAlign =3
                    Left =7867
                    Top =660
                    Width =870
                    Height =555
                    Name ="Label46"
                    Caption ="ID's   Scanned"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-15
                    Width =9195
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Cashiers Productivity Report"
                End
                Begin Label
                    TextAlign =3
                    Left =8760
                    Top =600
                    Width =780
                    Height =600
                    Name ="Label57"
                    Caption ="Return Qty "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9750
                    Top =600
                    Width =810
                    Height =600
                    Name ="Label58"
                    Caption ="Return Amount"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Height =225
                    ColumnWidth =840
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2400
                    Top =60
                    Width =1560
                    Height =225
                    ColumnWidth =1755
                    TabIndex =1
                    Name ="TotalTransactions"
                    ControlSource ="TotalTransactions"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3990
                    Top =60
                    Width =960
                    Height =225
                    ColumnWidth =1170
                    TabIndex =2
                    Name ="TotalSales"
                    ControlSource ="TotalSales"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    Width =900
                    Height =225
                    ColumnWidth =1830
                    TabIndex =3
                    Name ="AvgSale"
                    ControlSource ="AvgSale"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Top =60
                    Width =840
                    ColumnWidth =1305
                    TabIndex =4
                    Name ="TotalNoSales"
                    ControlSource ="TotalNoSales"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Width =780
                    ColumnWidth =1905
                    TabIndex =5
                    Name ="TotalCancelledSales"
                    ControlSource ="TotalCancelledSales"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =600
                    Height =225
                    ColumnWidth =390
                    TabIndex =6
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7927
                    Top =60
                    Width =780
                    TabIndex =7
                    Name ="Text47"
                    ControlSource ="TotalIDScanCount"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9600
                    Top =60
                    Width =1020
                    TabIndex =8
                    Name ="ReturnAmount"
                    ControlSource ="ReturnAmount"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8820
                    Top =60
                    Width =660
                    TabIndex =9
                    Name ="ReturnItems"
                    ControlSource ="ReturnItems"
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
            Height =1200
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    BorderWidth =2
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2220
                    Top =480
                    Width =1740
                    Height =300
                    FontSize =10
                    Name ="TotalTransactions Grand Total Sum"
                    ControlSource ="=Sum([TotalTransactions])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalTransactions_Grand_Total_Sum"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =480
                            Top =480
                            Width =1080
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label33"
                            Caption ="Grand Total"
                            FontName ="Arial Narrow"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3990
                    Top =480
                    Width =960
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="TotalSales Grand Total Sum"
                    ControlSource ="=Sum([TotalSales])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalSales_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6060
                    Top =480
                    Width =840
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="TotalNoSales Grand Total Sum"
                    ControlSource ="=Sum([TotalNoSales])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalNoSales_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7020
                    Top =480
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="TotalCancelledSales Grand Total Sum"
                    ControlSource ="=Sum([TotalCancelledSales])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalCancelledSales_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5040
                    Top =480
                    Width =960
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Text44"
                    ControlSource ="=Sum([TotalSales])/Sum([TotalTransactions])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =840
                    Top =60
                    Width =9780
                    Name ="Line45"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7927
                    Top =480
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text48"
                    ControlSource ="=Sum([TotalIDScanCount])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1920
                    Top =120
                    Width =2100
                    Height =660
                    Name ="Label51"
                    Caption ="Customer Count"
                    FontName ="Arial Narrow"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1920
                    LayoutCachedTop =120
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =780
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4020
                    Top =900
                    Width =4185
                    Height =300
                    Name ="Label52"
                    Caption ="* Note:  Sales do NOT include club invoices"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =480
                    Width =1020
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="Text55"
                    ControlSource ="=Sum([ReturnAmount])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    BorderWidth =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8760
                    Top =480
                    Width =660
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="Text56"
                    ControlSource ="=Sum([ReturnItems])"
                    FontName ="Arial Narrow"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD051CashiersProductivity.cls"
