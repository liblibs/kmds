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
    ItemSuffix =32
    Left =2700
    Top =75
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xbd5c9df189dce240
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
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            ControlSource ="PunchDate"
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
                    Left =-30
                    Width =10665
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Label47"
                    Caption ="Hours Worked By Cashier by Date "
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =15
                    Top =540
                    Width =10621
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =3180
                    Width =1125
                    Height =225
                    Name ="Label1"
                    Caption ="Business Date"
                End
                Begin Label
                    TextAlign =1
                    Left =1980
                    Width =1050
                    Height =225
                    Name ="Label3"
                    Caption ="Employee"
                End
                Begin Label
                    TextAlign =3
                    Left =5031
                    Width =360
                    Height =240
                    Name ="Label5"
                    Caption ="In"
                    LayoutCachedLeft =5031
                    LayoutCachedWidth =5391
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextAlign =3
                    Left =5886
                    Width =330
                    Height =225
                    Name ="Label7"
                    Caption ="Out"
                    LayoutCachedLeft =5886
                    LayoutCachedWidth =6216
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =7971
                    Width =1110
                    Height =225
                    Name ="Label9"
                    Caption ="Hours Worked"
                    LayoutCachedLeft =7971
                    LayoutCachedWidth =9081
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =9321
                    Width =1110
                    Height =225
                    Name ="lblWages"
                    Caption ="Wages"
                    LayoutCachedLeft =9321
                    LayoutCachedWidth =10431
                    LayoutCachedHeight =225
                End
                Begin Label
                    TextAlign =3
                    Left =6707
                    Width =360
                    Height =240
                    Name ="Label38"
                    Caption ="In"
                    LayoutCachedLeft =6707
                    LayoutCachedWidth =7067
                    LayoutCachedHeight =240
                End
                Begin Label
                    TextAlign =3
                    Left =7516
                    Width =330
                    Height =225
                    Name ="Label39"
                    Caption ="Out"
                    LayoutCachedLeft =7516
                    LayoutCachedWidth =7846
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =240
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1440
                    Width =720
                    Height =225
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2280
                    Width =1320
                    Height =225
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Width =1320
                    Height =225
                    TabIndex =2
                    Name ="Text29"
                    ControlSource ="PayrollID"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2400
                    Width =1800
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
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4724
                    Width =765
                    Height =225
                    TabIndex =1
                    Name ="Punch1"
                    ControlSource ="Punch1"
                    Format ="Short Time"

                    LayoutCachedLeft =4724
                    LayoutCachedWidth =5489
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5549
                    Width =765
                    Height =225
                    TabIndex =2
                    Name ="Punch2"
                    ControlSource ="Punch2"
                    Format ="Short Time"

                    LayoutCachedLeft =5549
                    LayoutCachedWidth =6314
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7874
                    Width =1305
                    Height =225
                    TabIndex =3
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

                    LayoutCachedLeft =7874
                    LayoutCachedWidth =9179
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
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9119
                    Height =225
                    TabIndex =4
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

                    LayoutCachedLeft =9119
                    LayoutCachedWidth =10559
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
                    Left =6400
                    Width =765
                    Height =225
                    TabIndex =5
                    Name ="Punch3"
                    ControlSource ="Punch3"
                    Format ="Short Time"

                    LayoutCachedLeft =6400
                    LayoutCachedWidth =7165
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7179
                    Width =765
                    Height =225
                    TabIndex =6
                    Name ="Punch4"
                    ControlSource ="Punch4"
                    Format ="Short Time"

                    LayoutCachedLeft =7179
                    LayoutCachedWidth =7944
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7320
                    Top =60
                    Width =1320
                    FontWeight =700
                    Name ="Text27"
                    ControlSource ="=Sum([HoursWorked])"
                    Format ="Standard"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =8700
                            Top =60
                            Width =690
                            Height =225
                            FontWeight =700
                            Name ="Label28"
                            Caption ="*"
                        End
                    End
                End
                Begin Line
                    Left =6000
                    Top =60
                    Width =2520
                    Name ="line5"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5820
                    Top =75
                    Width =1320
                    Height =225
                    TabIndex =1
                    Name ="Text30"
                    ControlSource ="PayrollID"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8955
                    Top =60
                    FontWeight =700
                    TabIndex =2
                    Name ="DTotalWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                    LayoutCachedLeft =8955
                    LayoutCachedTop =60
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =10455
                            Top =60
                            Width =180
                            Height =225
                            FontWeight =700
                            Name ="Label34"
                            Caption ="*"
                            LayoutCachedLeft =10455
                            LayoutCachedTop =60
                            LayoutCachedWidth =10635
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
                    Width =10560
                    FontSize =6
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
                    BackStyle =0
                    Left =45
                    Width =10560
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
            Height =2159
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
                    Width =9000
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
                    Width =7500
                    Height =1619
                    TabIndex =1
                    Name ="qryPunchHoursReport"
                    SourceObject ="Report.qryPunchHoursReport"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8955
                    Top =180
                    FontWeight =700
                    TabIndex =2
                    Name ="GTotalWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                    LayoutCachedLeft =8955
                    LayoutCachedTop =180
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "37 Hours Worked By Cashier By Date.cls"
