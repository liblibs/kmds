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
    Width =31680
    DatasheetFontHeight =10
    ItemSuffix =34
    RecSrcDt = Begin
        0x639d79f3158ae240
    End
    RecordSource ="qryEOChecksTender"
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
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            ControlSource ="Amount"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =3555
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="qryEOChecksTender"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =31680
                    BorderColor =12632256
                    Name ="Line31"
                End
            End
        End
        Begin PageHeader
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
                    TextAlign =3
                    Left =2880
                    Top =60
                    Width =1035
                    Height =300
                    Name ="Amount_Label"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3975
                    Top =60
                    Width =1560
                    Height =300
                    Name ="CashierName_Label"
                    Caption ="CashierName"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5595
                    Top =60
                    Width =1185
                    Height =300
                    Name ="CR_Label"
                    Caption ="CR"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6840
                    Top =60
                    Width =1185
                    Height =300
                    Name ="Invoice_Label"
                    Caption ="Invoice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =8085
                    Top =60
                    Width =2310
                    Height =300
                    Name ="Time_Label"
                    Caption ="Time"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =10455
                    Top =60
                    Width =1200
                    Height =300
                    Name ="CustomerID_Label"
                    Caption ="CustomerID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =11715
                    Top =60
                    Width =2310
                    Height =300
                    Name ="Customer_Label"
                    Caption ="Customer"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =14085
                    Top =60
                    Width =1395
                    Height =300
                    Name ="CheckNumber_Label"
                    Caption ="CheckNumber"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =15540
                    Top =60
                    Width =1080
                    Height =300
                    Name ="TranCount_Label"
                    Caption ="TranCount"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =31680
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =31680
                    Name ="Line33"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =390
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =60
                    Width =2760
                    Height =330
                    ColumnWidth =1065
                    FontSize =11
                    ForeColor =8388608
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =2880
                    Width =1035
                    Height =270
                    ColumnWidth =885
                    Name ="Amount"
                    ControlSource ="Amount"

                End
                Begin TextBox
                    Left =3975
                    Width =1560
                    Height =270
                    ColumnWidth =705
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
                Begin TextBox
                    Left =5595
                    Width =1185
                    Height =270
                    TabIndex =2
                    Name ="CR"
                    ControlSource ="CR"

                End
                Begin TextBox
                    Left =6840
                    Width =1185
                    Height =270
                    TabIndex =3
                    Name ="Invoice"
                    ControlSource ="Invoice"

                End
                Begin TextBox
                    Left =8085
                    Width =2310
                    Height =270
                    ColumnWidth =630
                    TabIndex =4
                    Name ="Time"
                    ControlSource ="Time"

                End
                Begin TextBox
                    Left =10455
                    Width =1200
                    Height =270
                    ColumnWidth =930
                    TabIndex =5
                    Name ="CustomerID"
                    ControlSource ="CustomerID"

                End
                Begin TextBox
                    Left =11715
                    Width =2310
                    Height =270
                    ColumnWidth =2070
                    TabIndex =6
                    Name ="Customer"
                    ControlSource ="Customer"

                End
                Begin TextBox
                    Left =14085
                    Width =1395
                    Height =270
                    ColumnWidth =1560
                    TabIndex =7
                    Name ="CheckNumber"
                    ControlSource ="CheckNumber"

                End
                Begin TextBox
                    Left =15540
                    Width =1080
                    Height =270
                    ColumnWidth =1140
                    TabIndex =8
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
                    Left =60
                    Width =31560
                    Name ="Text20"
                    ControlSource ="=\"Summary for \" & \"'CashierID' = \" & \" \" & [CashierID] & \" (\" & Count(*)"
                        " & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =270
                    Width =450
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label21"
                    Caption ="Sum"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2880
                    Top =270
                    Width =1035
                    Height =270
                    TabIndex =1
                    Name ="Sum Of Amount"
                    ControlSource ="=Sum([Amount])"
                    EventProcPrefix ="Sum_Of_Amount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =15540
                    Top =270
                    Width =1080
                    Height =270
                    TabIndex =2
                    Name ="Sum Of TranCount"
                    ControlSource ="=Sum([TranCount])"
                    EventProcPrefix ="Sum_Of_TranCount"

                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text29"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =26580
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =360
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    Left =2880
                    Width =1035
                    Height =300
                    FontSize =10
                    Name ="Amount Grand Total Sum"
                    ControlSource ="=Sum([Amount])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Amount_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
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
                    TextFontFamily =18
                    Left =15540
                    Width =1080
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="TranCount Grand Total Sum"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TranCount_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label26"
                    Caption ="Grand Total"
                End
            End
        End
    End
End
