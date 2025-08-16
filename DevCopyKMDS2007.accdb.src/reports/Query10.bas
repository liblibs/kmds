Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =124
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =33
    RecSrcDt = Begin
        0x1526ef2db2fae240
    End
    RecordSource ="Query10"
    Caption ="Query10"
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
            ControlSource ="CashSysID"
        End
        Begin BreakLevel
            ControlSource ="InvoHeadID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =0
                    Left =60
                    Top =60
                    Width =1500
                    Height =510
                    FontSize =20
                    Name ="Label27"
                    Caption ="Query10"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line30"
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =0
                    Left =60
                    Top =60
                    Width =1260
                    Height =300
                    Name ="CashSysID_Label"
                    Caption ="CashSysID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =1680
                    Top =60
                    Width =1500
                    Height =300
                    Name ="InvoHeadID_Label"
                    Caption ="InvoHeadID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3240
                    Top =60
                    Width =1200
                    Height =300
                    Name ="SetupDate_Label"
                    Caption ="SetupDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4080
                    Top =60
                    Width =900
                    Height =300
                    Name ="Code_Label"
                    Caption ="Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =5040
                    Top =60
                    Width =1740
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6840
                    Top =60
                    Width =780
                    Height =300
                    Name ="ShouldBe_Label"
                    Caption ="ShouldBe"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7680
                    Top =60
                    Width =780
                    Height =300
                    Name ="SoldAt_Label"
                    Caption ="SoldAt"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8520
                    Top =60
                    Width =780
                    Height =300
                    Name ="Diff_Label"
                    Caption ="Diff"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =9360
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =9360
                    Name ="Line32"
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
                    Width =2100
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
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
                    Left =900
                    Width =960
                    Height =270
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2220
                    Width =1800
                    Height =270
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4080
                    Width =900
                    Height =270
                    TabIndex =2
                    Name ="Code"
                    ControlSource ="Code"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5040
                    Width =1740
                    Height =270
                    ColumnWidth =3165
                    TabIndex =3
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6840
                    Width =780
                    Height =270
                    TabIndex =4
                    Name ="ShouldBe"
                    ControlSource ="ShouldBe"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7680
                    Width =780
                    Height =270
                    TabIndex =5
                    Name ="SoldAt"
                    ControlSource ="SoldAt"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8520
                    Width =780
                    Height =270
                    TabIndex =6
                    Name ="Diff"
                    ControlSource ="Diff"

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
                    IMESentenceMode =3
                    Left =60
                    Width =9240
                    Name ="Text16"
                    ControlSource ="=\"Summary for \" & \"'CashSysID' = \" & \" \" & [CashSysID] & \" (\" & Count(*)"
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
                    Name ="Label17"
                    Caption ="Sum"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =6840
                    Top =270
                    Width =780
                    Height =270
                    TabIndex =1
                    Name ="Sum Of ShouldBe"
                    ControlSource ="=Sum([ShouldBe])"
                    EventProcPrefix ="Sum_Of_ShouldBe"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =7680
                    Top =270
                    Width =780
                    Height =270
                    TabIndex =2
                    Name ="Sum Of SoldAt"
                    ControlSource ="=Sum([SoldAt])"
                    EventProcPrefix ="Sum_Of_SoldAt"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =8520
                    Top =270
                    Width =780
                    Height =270
                    TabIndex =3
                    Name ="Sum Of Diff"
                    ControlSource ="=Sum([Diff])"
                    EventProcPrefix ="Sum_Of_Diff"

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
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text28"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text29"
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
                    Left =6840
                    Width =780
                    Height =300
                    FontSize =10
                    Name ="ShouldBe Grand Total Sum"
                    ControlSource ="=Sum([ShouldBe])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="ShouldBe_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label19"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =7680
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="SoldAt Grand Total Sum"
                    ControlSource ="=Sum([SoldAt])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="SoldAt_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label22"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =8520
                    Width =780
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="Diff Grand Total Sum"
                    ControlSource ="=Sum([Diff])"
                    FontName ="Times New Roman"
                    EventProcPrefix ="Diff_Grand_Total_Sum"

                End
                Begin Label
                    Left =60
                    Width =1080
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label25"
                    Caption ="Grand Total"
                End
            End
        End
    End
End
