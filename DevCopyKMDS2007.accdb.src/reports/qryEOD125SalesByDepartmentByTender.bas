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
    ItemSuffix =70
    Left =570
    Top =1290
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x1dd61a054f15e440
    End
    RecordSource ="qryInvoLine5InvoheadByPayment"
    Caption ="qryEOD040SalesByDepartment1"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10680
                    Height =510
                    ColumnOrder =0
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
                    Left =585
                    Top =60
                    Width =9585
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Sales Totals by Sales Department By Tender Type"
                    LayoutCachedLeft =585
                    LayoutCachedTop =60
                    LayoutCachedWidth =10170
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line50"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =10800
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10800
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =600
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =1080
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3780
                    Top =120
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4800
                    Top =120
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6540
                    Top =120
                    Width =690
                    Height =300
                    Name ="D1_Label"
                    Caption =" Cash"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =120
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =7605
                    Top =120
                    Width =765
                    Height =300
                    Name ="D2_Label"
                    Caption ="Checks"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7605
                    LayoutCachedTop =120
                    LayoutCachedWidth =8370
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =8430
                    Top =120
                    Width =1200
                    Height =300
                    Name ="D3_Label"
                    Caption ="Charge"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8430
                    LayoutCachedTop =120
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =120
                    Width =855
                    Height =300
                    Name ="D4_Label"
                    Caption ="Other"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =120
                    LayoutCachedWidth =10755
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =150
                    Width =9720
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =120
                    Width =9720
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =450
                    Width =9720
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1080
                    Top =480
                    Width =9720
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =270
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =1140
                    Width =840
                    Height =270
                    Name ="Department"
                    ControlSource ="SalesDepartmentID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =2160
                    Width =1740
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    Left =3360
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="Count"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Width =1185
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="SumOfPrice"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="Cash"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="Checks"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Width =1185
                    Height =270
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="CreditCard"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1185
                    Height =270
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="Other"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"

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
            Height =1140
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3360
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([Count])"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([Cash])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D1_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([Checks])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([CreditCard])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D3_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Top =120
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=Sum([Other])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    FontName ="Arial Narrow"
                    EventProcPrefix ="D4_Grand_Total_Sum"

                End
                Begin Label
                    TextFontFamily =34
                    Top =120
                    Width =1155
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Tender Total"
                    FontName ="Arial Narrow"
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
                End
                Begin Line
                    Top =60
                    Width =10800
                    Name ="Line67"
                End
                Begin Label
                    Left =3120
                    Top =420
                    Width =7665
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label69"
                    Caption ="Note: If there are 2 payment type for an invoice then Charge is assumed for the "
                        "total invoice."
                    LayoutCachedLeft =3120
                    LayoutCachedTop =420
                    LayoutCachedWidth =10785
                    LayoutCachedHeight =690
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD125SalesByDepartmentByTender.cls"
