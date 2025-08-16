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
    Width =15720
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =1575
    Top =2490
    RecSrcDt = Begin
        0xe061b2511b2be340
    End
    RecordSource ="qryClubOrderGuidesForExport"
    Caption ="qryClubOrderGuidesForExport"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="LineID"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =225
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1800
                    Width =660
                    Height =225
                    ColumnWidth =735
                    Name ="LineID"
                    ControlSource ="LineID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2520
                    Width =660
                    Height =225
                    ColumnWidth =765
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3240
                    Width =2520
                    Height =225
                    ColumnWidth =2910
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5880
                    Width =660
                    Height =225
                    ColumnWidth =465
                    TabIndex =3
                    Name ="Par"
                    ControlSource ="Par"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10020
                    Width =1320
                    Height =225
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =11400
                    Width =1320
                    Height =225
                    TabIndex =5
                    Name ="OnSale"
                    ControlSource ="OnSale"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =12780
                    Width =1380
                    Height =225
                    TabIndex =6
                    Name ="AsOf"
                    ControlSource ="AsOf"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Width =1740
                    Height =225
                    ColumnWidth =2400
                    TabIndex =7
                    Name ="ClubName"
                    ControlSource ="ClubName"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =14280
                    Height =225
                    ColumnWidth =4575
                    TabIndex =8
                    Name ="Category"
                    ControlSource ="Category"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6720
                    Height =225
                    ColumnWidth =915
                    TabIndex =9
                    Name ="PkgSize"
                    ControlSource ="PkgSize"

                End
                Begin TextBox
                    DecimalPlaces =2
                    IMESentenceMode =3
                    Left =8220
                    Height =225
                    TabIndex =10
                    Name ="Price Per 1.25 Oz"
                    ControlSource ="PricePer125"
                    Format ="Standard"
                    EventProcPrefix ="Price_Per_1_25_Oz"

                End
            End
        End
    End
End
