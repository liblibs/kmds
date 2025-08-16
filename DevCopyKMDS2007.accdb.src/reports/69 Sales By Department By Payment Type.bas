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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =80
    Left =1665
    Top =1695
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xa1b8944995f3e540
    End
    RecordSource ="qryInvoLineByPayment"
    Caption ="qryEOD040SalesByDepartment1"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
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
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Line
            BorderColor =8388608
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            FontName ="Arial"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Subform
            OldBorderStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1320
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1500
                    Top =60
                    Width =7785
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption =" Sales By Sales Department by Payment Type"
                    LayoutCachedLeft =1500
                    LayoutCachedTop =60
                    LayoutCachedWidth =9285
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1260
                    Width =10680
                    BorderColor =12632256
                    Name ="Line53"
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
            Height =555
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =2760
                    Top =120
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =120
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =120
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3780
                    LayoutCachedTop =120
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =5520
                    Top =120
                    Width =690
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cash"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =6585
                    Top =120
                    Width =765
                    Height =300
                    Name ="D2_Label"
                    Caption ="Checks"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6585
                    LayoutCachedTop =120
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =7890
                    Top =120
                    Width =720
                    Height =300
                    Name ="D3_Label"
                    Caption ="Cards"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7890
                    LayoutCachedTop =120
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9930
                    Top =120
                    Width =690
                    Height =330
                    Name ="D4_Label"
                    Caption ="Other"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9930
                    LayoutCachedTop =120
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =450
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =120
                    Width =9600
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1080
                    Top =60
                    Width =9600
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =9600
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =540
                    Width =9600
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =120
                    Width =840
                    Height =270
                    Name ="Department"
                    ControlSource ="SalesDepartmentID"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =960
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Left =1140
                    Width =1740
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="Description"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    Left =2340
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="Count"

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3525
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3525
                    Width =1185
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="SumOfPrice"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =3525
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5040
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="Cash"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6225
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="Checks"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =6225
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7410
                    Width =1185
                    Height =270
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="CreditCard"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =7410
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Width =1065
                    Height =300
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="Other"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    Left =8820
                    Width =555
                    Height =270
                    TabIndex =8
                    Name ="Text77"
                    ControlSource ="=[CreditCard]/[SumOfPrice]"
                    Format ="Percent"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =9375
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =690
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
                    Width =4920
                    Height =330
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
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
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
                    Width =4920
                    Height =330
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
                    Width =10560
                    Height =330
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
                    Width =10560
                    Height =285
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
            CanGrow = NotDefault
            Height =1020
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2340
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([Count])"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                    LayoutCachedLeft =2340
                    LayoutCachedTop =120
                    LayoutCachedWidth =3525
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3525
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                    LayoutCachedLeft =3525
                    LayoutCachedTop =120
                    LayoutCachedWidth =4710
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =5040
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([Cash])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D1_Grand_Total_Sum"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =120
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6225
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([Checks])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D2_Grand_Total_Sum"

                    LayoutCachedLeft =6225
                    LayoutCachedTop =120
                    LayoutCachedWidth =7410
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7410
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([CreditCard])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D3_Grand_Total_Sum"

                    LayoutCachedLeft =7410
                    LayoutCachedTop =120
                    LayoutCachedWidth =8595
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Top =120
                    Width =1065
                    Height =330
                    FontWeight =700
                    TabIndex =5
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=Sum([Other])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D4_Grand_Total_Sum"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =120
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =450
                End
                Begin Label
                    Left =120
                    Top =120
                    Width =1620
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption =" Receipts Total"
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =390
                End
                Begin Label
                    Visible = NotDefault
                    Top =420
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    LayoutCachedTop =420
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =690
                End
                Begin Label
                    Left =2880
                    Top =480
                    Width =7665
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label69"
                    Caption ="Note: If there are 2 payment type for an invoice then Charge is assumed for the "
                        "total invoice."
                    LayoutCachedLeft =2880
                    LayoutCachedTop =480
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =750
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =3
                    Left =8820
                    Top =120
                    Width =555
                    Height =270
                    TabIndex =6
                    Name ="Text78"
                    ControlSource ="=Sum([CreditCard])/Sum([SumOfPrice])"
                    Format ="Percent"

                    LayoutCachedLeft =8820
                    LayoutCachedTop =120
                    LayoutCachedWidth =9375
                    LayoutCachedHeight =390
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =10620
                    Name ="Line79"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =10620
                End
            End
        End
    End
End
CodeBehindForm
' See "69 Sales By Department By Payment Type.cls"
