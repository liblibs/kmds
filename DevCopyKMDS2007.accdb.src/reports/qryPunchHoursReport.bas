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
    Width =8580
    DatasheetFontHeight =10
    ItemSuffix =30
    RecSrcDt = Begin
        0x5ec01ac954d8e240
    End
    RecordSource ="qryPunchHoursReport"
    Caption ="qryPunchHoursReport"
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
            ControlSource ="CashierID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1215
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =180
                    Width =3870
                    Height =510
                    FontSize =20
                    Name ="Label15"
                    Caption ="Recap"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =7140
                    BorderColor =12632256
                    Name ="Line18"
                End
                Begin Label
                    Left =2880
                    Top =840
                    Width =1560
                    Height =300
                    Name ="Label21"
                    Caption ="Employee"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4320
                    Top =840
                    Width =1080
                    Height =300
                    Name ="Label22"
                    Caption ="Payroll ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5940
                    Top =600
                    Width =1170
                    Height =540
                    Name ="Label23"
                    Caption ="Hours Worked"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =2940
                    Top =1200
                    Width =4140
                    Name ="Line24"
                End
                Begin TextBox
                    TextAlign =1
                    Left =180
                    Top =660
                    Width =5881
                    Height =300
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Format(Forms!PrintInventoryReport!txtDateRangeFrom,\"mm/dd/yyyy\""
                        ") & \" And \" & Format(Forms!PrintInventoryReport!txtDateRangeTo,\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
                Begin Label
                    TextAlign =3
                    Left =7680
                    Top =840
                    Width =675
                    Height =300
                    Name ="lblWages"
                    Caption ="Wages"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin PageHeader
            Visible = NotDefault
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =2760
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="CashierID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2880
                    Top =60
                    Width =1380
                    Height =300
                    Name ="CashierName_Label"
                    Caption ="CashierName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4320
                    Top =60
                    Width =1185
                    Height =300
                    Name ="PayrollID_Label"
                    Caption ="Payroll ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5745
                    Top =60
                    Width =1365
                    Height =300
                    Name ="HoursWorked_Label"
                    Caption ="HoursWorked"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =7140
                    Name ="Line19"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =7140
                    Name ="Line20"
                End
                Begin Label
                    TextAlign =3
                    Left =7650
                    Top =60
                    Width =675
                    Height =300
                    Name ="Label26"
                    Caption ="Wages"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5745
                    Width =1365
                    Height =270
                    Name ="HoursWorked"
                    ControlSource ="HoursWorked"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Width =1365
                    Height =270
                    TabIndex =1
                    Name ="Wages"
                    ControlSource ="Wages"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =255
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5745
                    Width =1365
                    Height =255
                    Name ="Sum Of HoursWorked"
                    ControlSource ="=Sum([HoursWorked])"
                    Format ="Standard"
                    EventProcPrefix ="Sum_Of_HoursWorked"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3060
                    Width =1200
                    Height =255
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4320
                    Width =1185
                    Height =255
                    TabIndex =2
                    Name ="PayrollID"
                    ControlSource ="PayrollID"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1980
                    Width =1020
                    Height =255
                    ColumnWidth =1020
                    TabIndex =3
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Width =1365
                    Height =255
                    TabIndex =4
                    Name ="CashierWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =300
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5745
                    Width =1365
                    Height =300
                    Name ="HoursWorked Grand Total Sum"
                    ControlSource ="=Sum([HoursWorked])"
                    Format ="Standard"
                    EventProcPrefix ="HoursWorked_Grand_Total_Sum"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6960
                    Width =1365
                    Height =300
                    TabIndex =1
                    Name ="TotalWages"
                    ControlSource ="=Sum([Wages])"
                    Format ="Standard"

                End
            End
        End
    End
End
CodeBehindForm
' See "qryPunchHoursReport.cls"
