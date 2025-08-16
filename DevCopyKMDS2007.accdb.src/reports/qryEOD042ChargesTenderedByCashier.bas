Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
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
    ItemSuffix =50
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x0d9ffd97ae0ee340
    End
    RecordSource ="qryEOChargesTendered"
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
            ControlSource ="CR"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            ControlSource ="Invoice"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1260
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
                    Left =2392
                    Top =60
                    Width =5940
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="EOD Charges Tendered by Cashier"
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
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =480
                    Top =60
                    Width =1020
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7020
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
                    Left =5940
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
                Begin Label
                    Left =4020
                    Top =60
                    Width =750
                    Height =300
                    Name ="Label38"
                    Caption ="TroutD"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4980
                    Top =60
                    Width =750
                    Height =300
                    Name ="Label40"
                    Caption ="Ticket"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =1260
                    Width =360
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="CR"
                    ControlSource ="CR"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Width =990
                            Height =300
                            Name ="Label42"
                            Caption ="Register #"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =540
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =60
                    Top =120
                    Width =540
                    Height =300
                    ColumnWidth =1065
                    FontSize =11
                    ForeColor =8388608
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =660
                    Top =120
                    Width =1872
                    Height =300
                    ColumnWidth =705
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashierName"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

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
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Width =1020
                    Height =225
                    ColumnWidth =885
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =2040
                    Width =1320
                    Height =225
                    TabIndex =1
                    Name ="Invoice"
                    ControlSource ="Invoice"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2040
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =3420
                    Width =765
                    Height =225
                    ColumnWidth =630
                    TabIndex =2
                    Name ="Time"
                    ControlSource ="Time"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    Left =6000
                    Width =525
                    Height =225
                    ColumnWidth =1140
                    TabIndex =3
                    Name ="TranCount"
                    ControlSource ="TranCount"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =4020
                    Width =765
                    Height =225
                    TabIndex =4
                    Name ="Text39"
                    ControlSource ="SalesDept"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =4980
                    Width =945
                    Height =225
                    TabIndex =5
                    Name ="Text41"
                    ControlSource ="SalesGroup"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10380
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Top =240
                    Width =1035
                    Height =225
                    Name ="Sum Of Amount"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_Amount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5460
                    Top =240
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Sum Of TranCount"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TranCount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2100
                    Top =240
                    Width =1560
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="CashierName"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =2100
                    Top =120
                    Width =5880
                    Name ="Line36"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6840
                    Top =180
                    Width =1035
                    Height =225
                    Name ="Text43"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5460
                    Top =180
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Text44"
                    ControlSource ="=Sum([TranCount])"

                End
                Begin Line
                    Left =2100
                    Top =60
                    Width =5880
                    Name ="Line46"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =2580
                    Top =120
                    Width =360
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text47"
                    ControlSource ="CR"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =1320
                            Top =120
                            Width =990
                            Height =300
                            Name ="Label48"
                            Caption ="Register #"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin Line
                    Left =2100
                    Top =120
                    Width =5880
                    Name ="Line49"
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
                    Left =6840
                    Top =60
                    Width =1035
                    Height =285
                    FontSize =10
                    Name ="Amount Grand Total Sum"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Amount_Grand_Total_Sum"

                End
                Begin Label
                    TextFontFamily =34
                    Left =1800
                    Top =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label23"
                    Caption ="Grand Total"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    Left =5400
                    Top =60
                    Width =1155
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text35"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Arial Narrow"

                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =7020
                    Top =60
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD042ChargesTenderedByCashier.cls"
