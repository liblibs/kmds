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
    Width =3600
    DatasheetFontHeight =10
    ItemSuffix =18
    RecSrcDt = Begin
        0x8bafd90b73abe240
    End
    RecordSource ="qryItemsWithUPCGlyph"
    Caption ="qryItemsWithUPCGlyph"
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
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1380
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2400
                    Top =540
                    Width =1140
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="Standard"

                End
                Begin TextBox
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =60
                    Top =540
                    Width =1500
                    Height =840
                    ColumnWidth =1740
                    FontSize =24
                    TabIndex =1
                    Name ="UPCG"
                    ControlSource ="UPCG"
                    FontName ="UPC-A"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =3540
                    Height =420
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Descr"
                    FontName ="Arial Narrow"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
