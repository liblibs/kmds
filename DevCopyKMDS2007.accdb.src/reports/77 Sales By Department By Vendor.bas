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
    ItemSuffix =82
    Left =11145
    Top =3375
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xbbbbe251ef1ce440
    End
    RecordSource ="qrySalesByDepartment"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1320
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =2295
                    Top =60
                    Width =6195
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption =" Sales Totals by Register Report"
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
            Height =600
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =660
                    Top =120
                    Width =1320
                    Height =300
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =660
                    LayoutCachedTop =120
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =3000
                    Top =120
                    Width =840
                    Height =300
                    Name ="TotalQ_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3000
                    LayoutCachedTop =120
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =3975
                    Top =120
                    Width =720
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Sales"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3975
                    LayoutCachedTop =120
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =5520
                    Top =120
                    Width =645
                    Height =300
                    Name ="D1_Label"
                    Caption =" Reg 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =6720
                    Top =120
                    Width =585
                    Height =300
                    Name ="D2_Label"
                    Caption ="Reg 2"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =120
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =7980
                    Top =120
                    Width =585
                    Height =300
                    Name ="D3_Label"
                    Caption ="Reg 3"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7980
                    LayoutCachedTop =120
                    LayoutCachedWidth =8565
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =9045
                    Top =120
                    Width =525
                    Height =330
                    Name ="D4_Label"
                    Caption =" Reg 4"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9045
                    LayoutCachedTop =120
                    LayoutCachedWidth =9570
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
                Begin Label
                    TextAlign =3
                    Left =10080
                    Top =120
                    Width =585
                    Height =330
                    Name ="Label77"
                    Caption =" Reg 5"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =10080
                    LayoutCachedTop =120
                    LayoutCachedWidth =10665
                    LayoutCachedHeight =450
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
            Height =270
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =720
                    Width =660
                    Height =270
                    Name ="Department"
                    ControlSource ="Department"

                    LayoutCachedLeft =720
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    Left =1440
                    Width =1800
                    Height =270
                    ColumnWidth =1800
                    TabIndex =1
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    Left =2940
                    Width =780
                    Height =270
                    TabIndex =2
                    Name ="TotalQ"
                    ControlSource ="TotalQ"

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3720
                    Width =1185
                    Height =270
                    FontWeight =700
                    TabIndex =3
                    Name ="TotalP"
                    ControlSource ="TotalP"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4995
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="D1"
                    ControlSource ="D1"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\015\012\""

                    LayoutCachedLeft =4995
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6180
                    Width =1185
                    Height =270
                    TabIndex =5
                    Name ="D2"
                    ControlSource ="D2"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =6180
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7365
                    Width =1185
                    Height =270
                    TabIndex =6
                    Name ="D3"
                    ControlSource ="D3"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =7365
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8550
                    Width =1065
                    Height =270
                    TabIndex =7
                    Name ="D4"
                    ControlSource ="D4"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =8550
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Width =1080
                    Height =270
                    TabIndex =8
                    Name ="D5"
                    ControlSource ="D5"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =270
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =540
            BreakLevel =1
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =180
                    Top =120
                    Width =330
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label3"
                    Caption ="*"
                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =510
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2940
                    Top =120
                    Width =780
                    Height =270
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="Sum_Of_TotalQ"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =120
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3720
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =1
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_TotalP"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =120
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4995
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D1"

                    LayoutCachedLeft =4995
                    LayoutCachedTop =120
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6180
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D2"

                    LayoutCachedLeft =6180
                    LayoutCachedTop =120
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7365
                    Top =120
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Sum Of D3"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D3"

                    LayoutCachedLeft =7365
                    LayoutCachedTop =120
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8550
                    Top =120
                    Width =1065
                    Height =270
                    TabIndex =5
                    Name ="Sum Of D4"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="Sum_Of_D4"

                    LayoutCachedLeft =8550
                    LayoutCachedTop =120
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =390
                End
                Begin Line
                    Left =60
                    Width =10620
                    Height =30
                    BorderColor =12632256
                    Name ="Line58"
                End
                Begin TextBox
                    Left =720
                    Top =120
                    Width =2160
                    Height =270
                    TabIndex =6
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"

                    LayoutCachedLeft =720
                    LayoutCachedTop =120
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9660
                    Top =120
                    Width =1020
                    Height =270
                    TabIndex =7
                    Name ="Text79"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =390
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
                    Left =2880
                    Top =60
                    Width =780
                    Height =270
                    Name ="Text69"
                    ControlSource ="=Sum([TotalQ])"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3660
                    Top =60
                    Width =1185
                    Height =270
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text70"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4935
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =4935
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6120
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Text72"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7305
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7305
                    Top =60
                    Width =1185
                    Height =270
                    TabIndex =4
                    Name ="Text73"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =7305
                    LayoutCachedTop =60
                    LayoutCachedWidth =8490
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8490
                    Top =60
                    Width =1065
                    Height =270
                    TabIndex =5
                    Name ="Text74"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =8490
                    LayoutCachedTop =60
                    LayoutCachedWidth =9555
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    Left =675
                    Top =60
                    Width =1785
                    Height =270
                    TabIndex =6
                    Name ="txtGroupLevelDesc"

                    LayoutCachedLeft =675
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =165
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label68"
                            Caption ="**"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =225
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =435
                    Top =60
                    Width =60
                    Height =270
                    TabIndex =7
                    Name ="txtGroupLevel"
                    ControlSource ="GroupLevel"

                    LayoutCachedLeft =435
                    LayoutCachedTop =60
                    LayoutCachedWidth =495
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9600
                    Top =60
                    Width =1020
                    Height =270
                    TabIndex =8
                    Name ="Text80"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =9600
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =330
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
                    Left =2940
                    Width =780
                    Height =300
                    FontWeight =700
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =3735
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([TotalP])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                    LayoutCachedLeft =3735
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =4980
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([D1])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D1_Grand_Total_Sum"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =6165
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6165
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =3
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([D2])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D2_Grand_Total_Sum"

                    LayoutCachedLeft =6165
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7350
                    Width =1185
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([D3])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D3_Grand_Total_Sum"

                    LayoutCachedLeft =7350
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =8535
                    Width =1065
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="D4 Grand Total Sum"
                    ControlSource ="=Sum([D4])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""
                    EventProcPrefix ="D4_Grand_Total_Sum"

                    LayoutCachedLeft =8535
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =300
                End
                Begin Label
                    Width =1320
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption =" Receipts Total"
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =270
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
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9660
                    Width =1020
                    Height =300
                    FontWeight =700
                    TabIndex =6
                    Name ="Text81"
                    ControlSource ="=Sum([D5])"
                    Format ="#,##0.00;-#,##0.00[Red];\"\";\"Null\""

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =300
                End
            End
        End
    End
End
CodeBehindForm
' See "77 Sales By Department By Vendor.cls"
