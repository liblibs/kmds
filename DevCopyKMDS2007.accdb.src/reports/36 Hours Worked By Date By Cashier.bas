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
    ItemSuffix =42
    Left =-510
    Top =1530
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x3dc4b6c689dce240
    End
    RecordSource ="qryPunchHoursReport"
    Caption ="PunchHoursReport"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="PunchDate"
        End
        Begin BreakLevel
            ControlSource ="CashierID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =840
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =-15
                    Width =10635
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Hours Worked By Date By Cashier"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =15
                    Top =540
                    Width =10516
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                    LayoutCachedLeft =15
                    LayoutCachedTop =540
                    LayoutCachedWidth =10531
                    LayoutCachedHeight =840
                End
            End
        End
        Begin PageHeader
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =600
                    Width =690
                    Height =225
                    Name ="Label1"
                    Caption ="BusDate"
                End
                Begin Label
                    TextAlign =1
                    Left =1920
                    Width =1050
                    Height =225
                    Name ="Label3"
                    Caption ="Employee"
                End
                Begin Label
                    TextAlign =3
                    Left =4965
                    Width =360
                    Height =240
                    Name ="Label5"
                    Caption ="In"
                    LayoutCachedLeft =4965
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextAlign =3
                    Left =5820
                    Width =330
                    Height =225
                    Name ="Label7"
                    Caption ="Out"
                    LayoutCachedLeft =5820
                    LayoutCachedWidth =6150
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =7905
                    Width =1110
                    Height =225
                    Name ="Label9"
                    Caption ="Hours Worked"
                    LayoutCachedLeft =7905
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =9158
                    Width =1110
                    Height =225
                    Name ="lblWages"
                    Caption ="Wages"
                    LayoutCachedLeft =9158
                    LayoutCachedWidth =10268
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =6641
                    Width =360
                    Height =240
                    Name ="Label38"
                    Caption ="In"
                    LayoutCachedLeft =6641
                    LayoutCachedWidth =7001
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextAlign =3
                    Left =7450
                    Width =330
                    Height =225
                    Name ="Label39"
                    Caption ="Out"
                    LayoutCachedLeft =7450
                    LayoutCachedWidth =7780
                    LayoutCachedHeight =225
                End
                Begin Line
                    Top =300
                    Width =8700
                    Name ="Line40"
                End
                Begin Label
                    TextAlign =1
                    Left =3480
                    Width =1050
                    Height =225
                    Name ="Label41"
                    Caption ="Payroll ID"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =2580
                    Height =225
                    ColumnWidth =1680
                    Name ="PunchDate"
                    ControlSource ="PunchDate"
                    Format ="ddd mm/dd/yyyy"
                    ConditionalFormat = Begin
                        0x01000000e8000000010000000100000000000000000000004300000001000000 ,
                        0x00000000c0c0c000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x46006f0072006d006100740028005b00500075006e0063006800440061007400 ,
                        0x65005d002c002200640064006400220029003d00220053006100740022002000 ,
                        0x4f007200200046006f0072006d006100740028005b00500075006e0063006800 ,
                        0x44006100740065005d002c002200640064006400220029003d00220053007500 ,
                        0x6e00220000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000c0c0c000420000004600 ,
                        0x6f0072006d006100740028005b00500075006e00630068004400610074006500 ,
                        0x5d002c002200640064006400220029003d002200530061007400220020004f00 ,
                        0x7200200046006f0072006d006100740028005b00500075006e00630068004400 ,
                        0x6100740065005d002c002200640064006400220029003d002200530075006e00 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =225
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1620
                    Width =420
                    Height =225
                    ColumnWidth =1020
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2040
                    Width =1320
                    Height =225
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4560
                    Width =765
                    Height =225
                    ColumnWidth =1995
                    TabIndex =2
                    Name ="Punch1"
                    ControlSource ="Punch1"
                    Format ="Short Time"

                    LayoutCachedLeft =4560
                    LayoutCachedWidth =5325
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5385
                    Width =765
                    Height =225
                    ColumnWidth =2100
                    TabIndex =3
                    Name ="Punch2"
                    ControlSource ="Punch2"
                    Format ="Short Time"

                    LayoutCachedLeft =5385
                    LayoutCachedWidth =6150
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7710
                    Width =1305
                    Height =225
                    TabIndex =4
                    Name ="HoursWorked"
                    ControlSource ="HoursWorked"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x01000000e8000000010000000100000000000000000000004300000001000000 ,
                        0x00000000c0c0c000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x46006f0072006d006100740028005b00500075006e0063006800440061007400 ,
                        0x65005d002c002200640064006400220029003d00220053006100740022002000 ,
                        0x4f007200200046006f0072006d006100740028005b00500075006e0063006800 ,
                        0x44006100740065005d002c002200640064006400220029003d00220053007500 ,
                        0x6e00220000000000
                    End

                    LayoutCachedLeft =7710
                    LayoutCachedWidth =9015
                    LayoutCachedHeight =225
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000c0c0c000420000004600 ,
                        0x6f0072006d006100740028005b00500075006e00630068004400610074006500 ,
                        0x5d002c002200640064006400220029003d002200530061007400220020004f00 ,
                        0x7200200046006f0072006d006100740028005b00500075006e00630068004400 ,
                        0x6100740065005d002c002200640064006400220029003d002200530075006e00 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3420
                    Width =1200
                    Height =225
                    TabIndex =5
                    Name ="Text29"
                    ControlSource ="PayrollID"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9195
                    Width =1215
                    Height =225
                    TabIndex =6
                    Name ="Wages"
                    ControlSource ="Wages"
                    Format ="Standard"
                    ConditionalFormat = Begin
                        0x01000000e8000000010000000100000000000000000000004300000001000000 ,
                        0x00000000c0c0c000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x46006f0072006d006100740028005b00500075006e0063006800440061007400 ,
                        0x65005d002c002200640064006400220029003d00220053006100740022002000 ,
                        0x4f007200200046006f0072006d006100740028005b00500075006e0063006800 ,
                        0x44006100740065005d002c002200640064006400220029003d00220053007500 ,
                        0x6e00220000000000
                    End

                    LayoutCachedLeft =9195
                    LayoutCachedWidth =10410
                    LayoutCachedHeight =225
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000c0c0c000420000004600 ,
                        0x6f0072006d006100740028005b00500075006e00630068004400610074006500 ,
                        0x5d002c002200640064006400220029003d002200530061007400220020004f00 ,
                        0x7200200046006f0072006d006100740028005b00500075006e00630068004400 ,
                        0x6100740065005d002c002200640064006400220029003d002200530075006e00 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Width =765
                    Height =225
                    TabIndex =7
                    Name ="Punch3"
                    ControlSource ="Punch3"
                    Format ="Short Time"

                    LayoutCachedLeft =6236
                    LayoutCachedWidth =7001
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7015
                    Width =765
                    Height =225
                    TabIndex =8
                    Name ="Punch4"
                    ControlSource ="Punch4"
                    Format ="Short Time"

                    LayoutCachedLeft =7015
                    LayoutCachedWidth =7780
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =300
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    FontWeight =700
                    Name ="Text27"
                    ControlSource ="=Sum([HoursWorked])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =8700
                            Top =60
                            Width =180
                            Height =225
                            FontWeight =700
                            Name ="Label28"
                            Caption ="*"
                        End
                    End
                End
                Begin Line
                    Left =7440
                    Top =60
                    Width =1200
                    Name ="line5"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8858
                    Top =60
                    FontWeight =700
                    TabIndex =1
                    Name ="DTotalWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                    LayoutCachedLeft =8858
                    LayoutCachedTop =60
                    LayoutCachedWidth =10298
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =10358
                            Top =60
                            Width =180
                            Height =225
                            FontWeight =700
                            Name ="Label34"
                            Caption ="*"
                            LayoutCachedLeft =10358
                            LayoutCachedTop =60
                            LayoutCachedWidth =10538
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =0
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
                    BackStyle =0
                    Left =5640
                    Top =300
                    Width =5040
                    Height =255
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
                    BackStyle =0
                    Left =60
                    Width =4260
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
                    BackStyle =0
                    Left =5640
                    Width =5040
                    Height =255
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
                    BackStyle =0
                    Top =300
                    Width =10485
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedTop =300
                    LayoutCachedWidth =10485
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    Left =45
                    Width =10485
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =45
                    LayoutCachedWidth =10530
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =3307
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7200
                    Top =180
                    FontWeight =700
                    Name ="Text17"
                    ControlSource ="=Sum([HoursWorked])"
                    Format ="Standard"

                End
                Begin Line
                    Left =120
                    Top =60
                    Width =9000
                    Name ="Line18"
                End
                Begin Line
                    Left =120
                    Width =8520
                    Name ="Line19"
                End
                Begin Label
                    TextFontFamily =34
                    Left =180
                    Top =180
                    Width =690
                    Height =225
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Total"
                End
                Begin Subform
                    Left =120
                    Top =540
                    Width =9645
                    Height =1619
                    TabIndex =1
                    Name ="qryPunchHoursReport"
                    SourceObject ="Report.qryPunchHoursReport"

                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =9765
                    LayoutCachedHeight =2159
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8858
                    Top =180
                    FontWeight =700
                    TabIndex =2
                    Name ="GTotalWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                    LayoutCachedLeft =8858
                    LayoutCachedTop =180
                    LayoutCachedWidth =10298
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "36 Hours Worked By Date By Cashier.cls"
