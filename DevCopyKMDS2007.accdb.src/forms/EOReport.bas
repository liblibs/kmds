Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4539
    DatasheetFontHeight =10
    ItemSuffix =11
    Top =600
    Bottom =600
    RecSrcDt = Begin
        0x9bf923d12ea8e240
    End
    RecordSource ="EOReport"
    Caption ="EOReport"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin FormHeader
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =1920
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =120
                    Width =2475
                    ColumnWidth =4365
                    Name ="ReportName"
                    ControlSource ="ReportName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =975
                            Height =240
                            Name ="Label1"
                            Caption ="ReportName"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =480
                    ColumnWidth =1680
                    TabIndex =1
                    Name ="NumberOfCopies"
                    ControlSource ="NumberOfCopies"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="Number of copies must be between 0 and 5"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =1290
                            Height =240
                            Name ="Label3"
                            Caption ="NumberOfCopies"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1935
                    Top =870
                    TabIndex =2
                    Name ="eMail"
                    ControlSource ="eMail"

                    LayoutCachedLeft =1935
                    LayoutCachedTop =870
                    LayoutCachedWidth =2195
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =600
                            Top =840
                            Width =1305
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label284"
                            Caption ="eMail"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =600
                            LayoutCachedTop =840
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1080
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1200
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Width =840
                            Height =240
                            Name ="Label5"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =540
                    TabIndex =1
                    Name ="upsize_ts"
                    ControlSource ="upsize_ts"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =540
                            Width =750
                            Height =240
                            Name ="Label9"
                            Caption ="upsize_ts"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =960
                    TabIndex =2
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =960
                            Width =975
                            Height =240
                            Name ="Label7"
                            Caption ="ChangeDate"
                        End
                    End
                End
            End
        End
    End
End
