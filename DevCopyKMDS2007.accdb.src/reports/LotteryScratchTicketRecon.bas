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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =71
    RecSrcDt = Begin
        0xc2a1f279b5dee240
    End
    RecordSource ="qryLotteryScratchReport"
    Caption ="Lottery Scratch Ticket Recon"
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
            GroupOn =4
            ControlSource ="BusinessDate"
        End
        Begin BreakLevel
            ControlSource ="GameNumber"
        End
        Begin BreakLevel
            ControlSource ="BusinessDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =5220
                    Height =510
                    FontSize =20
                    Name ="Label49"
                    Caption ="Lottery Scratch Ticket Recon"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line52"
                End
            End
        End
        Begin PageHeader
            Height =915
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =180
                    Top =540
                    Width =690
                    Height =300
                    Name ="GameNumber_Label"
                    Caption ="Ticket "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4140
                    Top =540
                    Width =660
                    Height =300
                    Name ="BeginMeter_Label"
                    Caption ="Begin "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5130
                    Top =540
                    Width =450
                    Height =300
                    Name ="EndMeter_Label"
                    Caption ="End"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5925
                    Top =540
                    Width =435
                    Height =300
                    Name ="AddMeter_Label"
                    Caption ="Add"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =7395
                    Top =540
                    Width =825
                    Height =300
                    Name ="AmountSold_Label"
                    Caption ="Register"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6570
                    Top =540
                    Width =630
                    Height =300
                    Name ="PhysicalAmount_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9495
                    Top =540
                    Width =900
                    Height =300
                    Name ="Variance_Label"
                    Caption ="Variance"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =480
                    Width =10800
                    Name ="Line53"
                End
                Begin Line
                    BorderWidth =2
                    Top =900
                    Width =10800
                    Name ="Line54"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    IMESentenceMode =3
                    Top =60
                    Width =2760
                    Height =330
                    FontSize =11
                    ForeColor =8388608
                    Name ="BusinessDate by Month"
                    ControlSource ="BusinessDate"
                    Format ="Long Date"
                    FontName ="Times New Roman"
                    EventProcPrefix ="BusinessDate_by_Month"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =495
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Top =120
                    Width =360
                    Height =225
                    Name ="GameNumber"
                    ControlSource ="GameNumber"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =360
                    Top =120
                    Width =780
                    Height =225
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4080
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="BeginMeter"
                    ControlSource ="BeginMeter"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="EndMeter"
                    ControlSource ="EndMeter"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="AddMeter"
                    ControlSource ="AddMeter"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7500
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="AmountSold"
                    ControlSource ="AmountSold"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="PhysicalAmount"
                    ControlSource ="PhysicalAmount"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =120
                    Width =720
                    Height =225
                    TabIndex =7
                    Name ="Variance"
                    ControlSource ="Variance"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =120
                    Width =975
                    Height =225
                    TabIndex =8
                    Name ="VarianceDollars"
                    ControlSource ="VarianceDollars"
                    Format ="#,###.00;-#,###.00[Red];\"\""

                End
                Begin Line
                    Visible = NotDefault
                    Top =480
                    Width =10440
                    Name ="Line5"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1200
                    Top =120
                    Width =2880
                    ColumnWidth =3150
                    TabIndex =9
                    Name ="Description"
                    ControlSource ="Description"

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
                    Width =10680
                    Name ="Text26"
                    ControlSource ="=\"Summary for \" & \"'BusinessDate' = \" & \" \" & [BusinessDate] & \" (\" & Co"
                        "unt(*) & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                End
                Begin Label
                    FontItalic = NotDefault
                    Left =60
                    Top =270
                    Width =450
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label27"
                    Caption ="Sum"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4080
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text55"
                    ControlSource ="=Sum([BeginMeter])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text56"
                    ControlSource ="=Sum([EndMeter])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Text57"
                    ControlSource ="=Sum([AddMeter])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7500
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text58"
                    ControlSource ="=Sum([AmountSold])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text59"
                    ControlSource ="=Sum([PhysicalAmount])"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =300
                    Width =720
                    Height =225
                    TabIndex =6
                    Name ="Text60"
                    ControlSource ="=Sum([Variance])"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =300
                    Width =975
                    Height =225
                    TabIndex =7
                    Name ="Text61"
                    ControlSource ="=Sum([VarianceDollars])"
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
            Visible = NotDefault
            Height =285
            Name ="ReportFooter"
            Begin
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label32"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label35"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label41"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label44"
                    Caption ="Grand Total"
                End
                Begin Label
                    Left =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label47"
                    Caption ="Grand Total"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4080
                    Top =60
                    Width =720
                    Height =225
                    Name ="Text62"
                    ControlSource ="=Sum([BeginMeter])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text63"
                    ControlSource ="=Sum([EndMeter])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5640
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text64"
                    ControlSource ="=Sum([AddMeter])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7500
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =3
                    Name ="Text65"
                    ControlSource ="=Sum([AmountSold])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =4
                    Name ="Text66"
                    ControlSource ="=Sum([PhysicalAmount])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8640
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =5
                    Name ="Text67"
                    ControlSource ="=Sum([Variance])"
                    Format ="#,##0;-#,##0[Red];\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Width =975
                    Height =225
                    TabIndex =6
                    Name ="Text68"
                    ControlSource ="=Sum([VarianceDollars])"
                    Format ="#,###.00;-#,###.00[Red];\"\""

                End
            End
        End
    End
End
CodeBehindForm
' See "LotteryScratchTicketRecon.cls"
