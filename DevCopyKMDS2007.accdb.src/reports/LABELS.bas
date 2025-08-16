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
    GridX =20
    GridY =24
    Width =3780
    DatasheetFontHeight =10
    ItemSuffix =10
    RecSrcDt = Begin
        0xd2323e2c4aade240
    End
    RecordSource ="qryItemLabelsWithUPCGlyph"
    OnOpen ="=LabelSetup()"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            AutoLabel = NotDefault
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            FontWeight =300
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin BreakLevel
            ControlSource ="Descr"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            OnFormat ="=LabelInitialize()"
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            OnFormat ="[Event Procedure]"
            OnPrint ="=LabelLayout(Reports!LabelS)"
            Name ="Detail"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =72
                    Top =60
                    Width =3168
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Stephen Kelly"
                End
                Begin Label
                    TextFontFamily =34
                    Left =72
                    Top =360
                    Width =3168
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label6"
                    Caption ="11605 SW 31st Ct."
                End
                Begin Label
                    TextFontFamily =34
                    Left =72
                    Top =660
                    Width =3168
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Label7"
                    Caption ="Portland, OR   97219"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =75
                    Top =960
                    Width =3630
                    Height =315
                    FontSize =10
                    FontWeight =700
                    Name ="Label8"
                    Caption ="(503) 452-1066   skelly@comcast.net"
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
CodeBehindForm
' See "LABELS.cls"
