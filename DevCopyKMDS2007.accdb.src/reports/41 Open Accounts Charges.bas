Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =39
    Left =525
    Top =1230
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x5157b38d8b70e340
    End
    RecordSource ="qryOpenAccountsTendered"
    Caption ="qryEOChecksTender"
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
            ControlSource ="CustomerID"
        End
        Begin BreakLevel
            ControlSource ="Amount"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1500
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10620
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =60
                    Width =10725
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="EOD Open Accounts by Customer"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10740
                    BorderColor =12632256
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10740
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Top =1200
                    Width =10621
                    Height =300
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =5400
                    Top =60
                    Width =1020
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =60
                    Width =810
                    Height =300
                    Name ="Amount_Label"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2220
                    Top =60
                    Width =360
                    Height =300
                    Name ="CR_Label"
                    Caption ="CR"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2640
                    Top =60
                    Width =750
                    Height =300
                    Name ="Invoice_Label"
                    Caption ="Invoice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3420
                    Top =60
                    Width =540
                    Height =300
                    Name ="Time_Label"
                    Caption ="Time"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =60
                    Top =60
                    Width =1200
                    Height =300
                    Name ="CustomerID_Label"
                    Caption ="Customer"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7200
                    Top =60
                    Width =1575
                    Height =300
                    Name ="CheckNumber_Label"
                    Caption ="CheckNumber"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8820
                    Top =60
                    Width =630
                    Height =300
                    Name ="TranCount_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10800
                    Name ="Line33"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =540
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =60
                    Top =120
                    Width =1260
                    Height =300
                    ColumnWidth =1065
                    FontSize =11
                    ForeColor =8388608
                    Name ="CashierID"
                    ControlSource ="CustomerID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1440
                    Top =120
                    Width =1872
                    Height =300
                    ColumnWidth =705
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashierName"
                    ControlSource ="Customer"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Width =1035
                    Height =225
                    ColumnWidth =885
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"

                End
                Begin TextBox
                    Left =1080
                    Width =360
                    Height =225
                    TabIndex =1
                    Name ="CR"
                    ControlSource ="CR"

                End
                Begin TextBox
                    Left =1500
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Invoice"
                    ControlSource ="Invoice"

                End
                Begin TextBox
                    Left =2520
                    Width =1980
                    Height =225
                    ColumnWidth =630
                    TabIndex =3
                    Name ="Time"
                    ControlSource ="SetupDate"

                End
                Begin TextBox
                    TextAlign =3
                    Left =4320
                    Width =600
                    Height =225
                    ColumnWidth =930
                    TabIndex =4
                    Name ="CustomerID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    Left =4980
                    Width =2310
                    Height =225
                    ColumnWidth =2070
                    TabIndex =5
                    Name ="Customer"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    Left =7380
                    Width =1395
                    Height =225
                    ColumnWidth =1560
                    TabIndex =6
                    Name ="CheckNumber"
                    ControlSource ="CheckNumber"

                End
                Begin TextBox
                    TextAlign =3
                    Left =8880
                    Width =540
                    Height =225
                    ColumnWidth =1140
                    TabIndex =7
                    Name ="TranCount"
                    ControlSource ="TranCount"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Top =240
                    Width =1035
                    Height =225
                    Name ="Sum Of Amount"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    EventProcPrefix ="Sum_Of_Amount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =8340
                    Top =240
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Sum Of TranCount"
                    ControlSource ="=Sum([TranCount])"
                    EventProcPrefix ="Sum_Of_TranCount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =6660
                    Top =240
                    Width =1560
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="CashierName"

                End
                Begin Line
                    Left =6660
                    Top =120
                    Width =4080
                    Name ="Line36"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =3180
                    Top =180
                    Width =1260
                    Height =300
                    FontSize =11
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text37"
                    ControlSource ="CustomerID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =4560
                    Top =180
                    Width =1872
                    Height =300
                    FontSize =11
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text38"
                    ControlSource ="Customer"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin PageFooter
            Height =660
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =360
                    Width =5040
                    Height =300
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
                    Left =5760
                    Top =360
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line60"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
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
                    Left =5760
                    Top =60
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
                    Left =120
                    Top =360
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
                    Left =120
                    Top =60
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
            KeepTogether = NotDefault
            Height =480
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =18
                    Left =9720
                    Top =60
                    Width =1035
                    Height =285
                    FontSize =10
                    Name ="Amount Grand Total Sum"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Amount_Grand_Total_Sum"

                End
                Begin Label
                    Left =6360
                    Top =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label23"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =18
                    Left =8280
                    Top =60
                    Width =1155
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text35"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Visible = NotDefault
                    Left =1020
                    Top =120
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                End
            End
        End
    End
End
CodeBehindForm
' See "41 Open Accounts Charges.cls"
