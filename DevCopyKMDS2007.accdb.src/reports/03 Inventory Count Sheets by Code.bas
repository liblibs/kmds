Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =86
    Left =585
    Top =1320
    OrderBy ="ID"
    RecSrcDt = Begin
        0x8ab255c558abe240
    End
    RecordSource ="qryInventoryByItemID"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =10635
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Count Sheets By Code"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =840
                    Width =10621
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboItemIDBegin & \" and \" & Forms!Pr"
                        "intInventoryReport!cboItemIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =4890
                    Width =1215
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="On the shelf"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =300
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Width =10710
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6270
                    Width =930
                    Height =300
                    Name ="Label79"
                    Caption ="Reserved"
                End
                Begin Label
                    TextAlign =3
                    Left =7305
                    Width =945
                    Height =300
                    Name ="Label80"
                    Caption ="Total"
                End
                Begin Label
                    TextAlign =3
                    Left =8460
                    Width =2115
                    Height =300
                    Name ="Label82"
                    Caption ="Physical count"
                End
                Begin Label
                    TextAlign =2
                    Left =1080
                    Width =960
                    Height =300
                    Name ="Label85"
                    Caption ="Item"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =660
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =540
                    Top =180
                    Width =1320
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedLeft =540
                    LayoutCachedTop =180
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =405
                End
                Begin TextBox
                    Left =1980
                    Top =180
                    Width =2460
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    Left =5385
                    Top =180
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="OnTheShelf"
                    ControlSource ="OnTheShelf"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4500
                    Top =180
                    Width =840
                    TabIndex =3
                    Name ="SizeD"
                    ControlSource ="SizeD"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6270
                    Top =180
                    Width =930
                    TabIndex =4
                    Name ="OnReserve"
                    ControlSource ="OnReserve"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7320
                    Top =180
                    Width =930
                    TabIndex =5
                    Name ="OnHand"
                    ControlSource ="=[OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve])"

                End
                Begin Line
                    Left =9000
                    Top =300
                    Width =1620
                    Name ="Line83"
                End
                Begin Line
                    Top =600
                    Width =10620
                    Name ="line5"
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
                    Left =5700
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
                    Left =60
                    Width =4380
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
                    Left =5700
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
                    Top =300
                    Width =10680
                    Height =255
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
                    Width =10680
                    Height =255
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "03 Inventory Count Sheets by Code.cls"
