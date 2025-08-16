Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3300
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =630
    Top =1800
    RecSrcDt = Begin
        0xb1e6d922b2c0e240
    End
    RecordSource ="qryClubOrderSizeTotals"
    Caption ="qryClubOrderSizeTotals"
    DatasheetFontName ="Trebuchet MS"
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
            Height =615
            Name ="ReportHeader"
            Begin
                Begin Label
                    Left =180
                    Top =300
                    Width =480
                    Height =240
                    ForeColor =0
                    Name ="Label16"
                    Caption ="Size "
                End
                Begin Label
                    TextAlign =3
                    Left =945
                    Top =300
                    Width =675
                    Height =240
                    ForeColor =0
                    Name ="Label17"
                    Caption ="Bottles"
                End
                Begin Label
                    TextAlign =3
                    Left =1890
                    Width =630
                    Height =540
                    ForeColor =0
                    Name ="Label18"
                    Caption ="Full Cases"
                End
                Begin Label
                    TextAlign =3
                    Left =2610
                    Top =240
                    Width =630
                    Height =300
                    ForeColor =0
                    Name ="Label19"
                    Caption ="Break"
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Top =600
                    Width =3120
                    Name ="Line11"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =180
                    Width =540
                    Height =225
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    Left =1740
                    Width =780
                    Height =225
                    TabIndex =1
                    Name ="FirstOfCaseQty"
                    ControlSource ="Full"

                End
                Begin TextBox
                    TextAlign =3
                    Left =810
                    Width =840
                    Height =225
                    TabIndex =2
                    Name ="SumOfQuantityShipped"
                    ControlSource ="QS"

                End
                Begin TextBox
                    TextAlign =3
                    Left =2580
                    Width =660
                    Height =225
                    TabIndex =3
                    Name ="Text20"
                    ControlSource ="Broken"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =390
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =810
                    Top =120
                    Width =840
                    Height =270
                    Name ="Text26"
                    ControlSource ="=Sum([QS])"

                End
                Begin TextBox
                    TextAlign =3
                    Left =1740
                    Top =120
                    Width =780
                    Height =270
                    TabIndex =1
                    Name ="Text27"
                    ControlSource ="=Sum([Full])"

                End
                Begin TextBox
                    TextAlign =3
                    Left =2640
                    Top =120
                    Width =600
                    Height =270
                    TabIndex =2
                    Name ="Text28"
                    ControlSource ="=Sum([Broken])"

                End
                Begin Line
                    Left =780
                    Top =60
                    Width =2460
                    Name ="Line29"
                End
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =690
                    Height =300
                    ForeColor =0
                    Name ="Label13"
                    Caption ="Totals "
                End
            End
        End
    End
End
