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
    ItemSuffix =78
    Left =1050
    Top =1860
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x06702adb2cfbe240
    End
    RecordSource ="qrySalesByDepartmentByCategory"
    Caption ="qryEOD040SalesByDepartment1"
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
            ControlSource ="GroupLevel"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Group"
        End
        Begin BreakLevel
            ControlSource ="Department"
        End
        Begin BreakLevel
            ControlSource ="SubCategoryDescr"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1320
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10560
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption =" Sales Totals by Department by Category"
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
                    Left =6585
                    Top =120
                    Width =645
                    Height =300
                    Name ="D1_Label"
                    Caption =" Reg 1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7785
                    Top =120
                    Width =585
                    Height =300
                    Name ="D2_Label"
                    Caption ="Reg 2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9045
                    Top =120
                    Width =585
                    Height =300
                    Name ="D3_Label"
                    Caption ="Reg 3"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10110
                    Top =120
                    Width =525
                    Height =330
                    Name ="D4_Label"
                    Caption =" Reg 4"
                    Tag ="DetachedLabel"
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
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =300
                    Width =360
                    Height =270
                    Name ="Department"
                    ControlSource ="Department"

                End
                Begin TextBox
                    Left =720
                    Width =2760
                    Height =270
                    TabIndex =1
                    Name ="DepartmentDescription"
                    ControlSource ="=[DepartmentDesc] & \"-\" & [SubCategoryDescr]"

                End
                Begin TextBox
                    TextAlign =3
                    Left =3360
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="TotalQ"

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
                    ControlSource ="TotalP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="D1"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="D2"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Width =1185
                    Height =270
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="D3"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1065
                    Height =300
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="D4"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =660
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =240
                    Top =150
                    Width =450
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="*"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3360
                    Top =120
                    Width =1185
                    Height =270
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="Sum_Of_TotalQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =1
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_TotalP"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D2"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Sum Of D3"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D3"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Top =120
                    Width =1065
                    Height =300
                    TabIndex =5
                    Name ="Sum Of D4"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D4"

                End
                Begin Line
                    Left =60
                    Width =10620
                    Height =30
                    BorderColor =12632256
                    Name ="Line58"
                End
                Begin TextBox
                    Left =780
                    Top =150
                    Width =2460
                    TabIndex =6
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3300
                    Top =60
                    Width =1185
                    Height =270
                    Name ="Text69"
                    ControlSource ="=Sum([TotalQ])"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4485
                    Top =60
                    Width =1185
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6000
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7185
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Text72"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8370
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Text73"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9555
                    Top =60
                    Width =1065
                    Height =300
                    TabIndex =5
                    Name ="Text74"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                End
                Begin TextBox
                    Left =735
                    Top =120
                    Width =2100
                    TabIndex =6
                    Name ="txtGroupLevelDesc"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =120
                            Top =120
                            Width =285
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label68"
                            Caption ="**"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =495
                    Top =120
                    Width =180
                    TabIndex =7
                    Name ="txtGroupLevel"
                    ControlSource ="GroupLevel"

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
            Height =570
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =3360
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4545
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6060
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D1_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7245
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8430
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D3_Grand_Total_Sum"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9615
                    Width =1065
                    Height =330
                    FontWeight =700
                    TabIndex =5
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D4_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1620
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption =" Receipts Total"
                End
                Begin Label
                    Visible = NotDefault
                    Top =300
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
' See "42 Sales By Department By Category.cls"
