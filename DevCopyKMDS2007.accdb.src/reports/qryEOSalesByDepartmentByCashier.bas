Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =2
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =63
    RecSrcDt = Begin
        0xd67f26710da2e240
    End
    RecordSource ="qryEOSalesByDepartmentByCashier"
    Caption ="qryEOSalesByDepartmentByCashier"
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
            ControlSource ="Group"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =705
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =4530
                    Height =285
                    FontSize =20
                    Name ="Label57"
                    Caption ="qryEOSalesByDepartmentByCashier"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line60"
                End
            End
        End
        Begin PageHeader
            Height =420
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1320
                    Height =285
                    Name ="Group_Label"
                    Caption ="Group"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1440
                    Top =60
                    Width =960
                    Height =285
                    Name ="GroupDesc_Label"
                    Caption ="GroupDesc"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2460
                    Top =60
                    Width =720
                    Height =285
                    Name ="Department_Label"
                    Caption ="Department"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3240
                    Top =60
                    Width =1080
                    Height =285
                    Name ="DepartmentDesc_Label"
                    Caption ="DepartmentDesc"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4380
                    Top =60
                    Width =540
                    Height =285
                    Name ="TotalQ_Label"
                    Caption ="TotalQ"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4980
                    Top =60
                    Width =540
                    Height =285
                    Name ="TotalP_Label"
                    Caption ="TotalP"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5580
                    Top =60
                    Width =540
                    Height =285
                    Name ="Q1_Label"
                    Caption ="Q1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6180
                    Top =60
                    Width =540
                    Height =285
                    Name ="D1_Label"
                    Caption ="D1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6780
                    Top =60
                    Width =540
                    Height =285
                    Name ="Q2_Label"
                    Caption ="Q2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7380
                    Top =60
                    Width =540
                    Height =285
                    Name ="D2_Label"
                    Caption ="D2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7980
                    Top =60
                    Width =540
                    Height =285
                    Name ="Q3_Label"
                    Caption ="Q3"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8580
                    Top =60
                    Width =540
                    Height =285
                    Name ="D3_Label"
                    Caption ="D3"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9180
                    Top =60
                    Width =540
                    Height =285
                    Name ="Q4_Label"
                    Caption ="Q4"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9780
                    Top =60
                    Width =960
                    Height =285
                    Name ="D4_Label"
                    Caption ="D4"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line61"
                End
                Begin Line
                    BorderWidth =2
                    Top =405
                    Width =10800
                    Name ="Line62"
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
                    IMESentenceMode =3
                    Left =60
                    Width =1320
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="Group"
                    ControlSource ="Group"
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
                    IMESentenceMode =3
                    Left =1440
                    Width =960
                    Height =270
                    Name ="GroupDesc"
                    ControlSource ="GroupDesc"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2460
                    Width =720
                    Height =270
                    TabIndex =1
                    Name ="Department"
                    ControlSource ="Department"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3240
                    Width =1080
                    Height =270
                    ColumnWidth =1635
                    TabIndex =2
                    Name ="DepartmentDesc"
                    ControlSource ="DepartmentDesc"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4380
                    Width =540
                    Height =270
                    TabIndex =3
                    Name ="TotalQ"
                    ControlSource ="TotalQ"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4980
                    Width =540
                    Height =270
                    TabIndex =4
                    Name ="TotalP"
                    ControlSource ="TotalP"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5580
                    Width =540
                    Height =270
                    TabIndex =5
                    Name ="Q1"
                    ControlSource ="Q1"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Width =540
                    Height =270
                    TabIndex =6
                    Name ="D1"
                    ControlSource ="D1"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6780
                    Width =540
                    Height =270
                    TabIndex =7
                    Name ="Q2"
                    ControlSource ="Q2"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7380
                    Width =540
                    Height =270
                    TabIndex =8
                    Name ="D2"
                    ControlSource ="D2"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7980
                    Width =540
                    Height =270
                    TabIndex =9
                    Name ="Q3"
                    ControlSource ="Q3"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8580
                    Width =540
                    Height =270
                    TabIndex =10
                    Name ="D3"
                    ControlSource ="D3"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9180
                    Width =540
                    Height =270
                    TabIndex =11
                    Name ="Q4"
                    ControlSource ="Q4"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9780
                    Width =960
                    Height =270
                    TabIndex =12
                    Name ="D4"
                    ControlSource ="D4"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =630
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =60
                    Width =10680
                    Name ="Text28"
                    ControlSource ="=\"Summary for \" & \"'Group' = \" & \" \" & [Group] & \" (\" & Count(*) & \" \""
                        " & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =285
                    Width =510
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Sum"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =4380
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =1
                    Name ="Sum Of TotalQ"
                    ControlSource ="=Sum([TotalQ])"
                    EventProcPrefix ="Sum_Of_TotalQ"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =4980
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =2
                    Name ="Sum Of TotalP"
                    ControlSource ="=Sum([TotalP])"
                    EventProcPrefix ="Sum_Of_TotalP"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =5580
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =3
                    Name ="Sum Of Q1"
                    ControlSource ="=Sum([Q1])"
                    EventProcPrefix ="Sum_Of_Q1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =6180
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =4
                    Name ="Sum Of D1"
                    ControlSource ="=Sum([D1])"
                    EventProcPrefix ="Sum_Of_D1"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =6780
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =5
                    Name ="Sum Of Q2"
                    ControlSource ="=Sum([Q2])"
                    EventProcPrefix ="Sum_Of_Q2"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =7380
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =6
                    Name ="Sum Of D2"
                    ControlSource ="=Sum([D2])"
                    EventProcPrefix ="Sum_Of_D2"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =7980
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =7
                    Name ="Sum Of Q3"
                    ControlSource ="=Sum([Q3])"
                    EventProcPrefix ="Sum_Of_Q3"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =8580
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =8
                    Name ="Sum Of D3"
                    ControlSource ="=Sum([D3])"
                    EventProcPrefix ="Sum_Of_D3"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =9180
                    Top =285
                    Width =540
                    Height =285
                    TabIndex =9
                    Name ="Sum Of Q4"
                    ControlSource ="=Sum([Q4])"
                    EventProcPrefix ="Sum_Of_Q4"

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
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text58"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5700
                    Top =240
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text59"
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
                    IMESentenceMode =3
                    Left =4380
                    Width =540
                    Height =300
                    FontSize =10
                    Name ="TotalQ Grand Total Sum"
                    ControlSource ="=Sum([TotalQ])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TotalQ_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4980
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="TotalP Grand Total Sum"
                    ControlSource ="=Sum([TotalP])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TotalP_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label34"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Q1 Grand Total Sum"
                    ControlSource ="=Sum([Q1])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Q1_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label37"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6180
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="D1 Grand Total Sum"
                    ControlSource ="=Sum([D1])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="D1_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label40"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6780
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="Q2 Grand Total Sum"
                    ControlSource ="=Sum([Q2])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Q2_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label43"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7380
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="D2 Grand Total Sum"
                    ControlSource ="=Sum([D2])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="D2_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label46"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7980
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="Q3 Grand Total Sum"
                    ControlSource ="=Sum([Q3])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Q3_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8580
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="D3 Grand Total Sum"
                    ControlSource ="=Sum([D3])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="D3_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label52"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9180
                    Width =540
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Q4 Grand Total Sum"
                    ControlSource ="=Sum([Q4])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Q4_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1650
                    Height =285
                    FontSize =10
                    ForeColor =0
                    Name ="Label55"
                    Caption ="Grand Total"
                End
            End
        End
    End
End
