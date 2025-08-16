Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4539
    DatasheetFontHeight =10
    ItemSuffix =8
    RecSrcDt = Begin
        0xd14a5b99efa8e240
    End
    RecordSource ="InventoryReportOrderBys"
    Caption ="InventoryReportOrderBys"
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =840
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =60
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =60
                            Width =240
                            Height =240
                            Name ="Label1"
                            Caption ="ID"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =540
                    Width =2475
                    ColumnWidth =3840
                    TabIndex =1
                    Name ="ReportName"
                    ControlSource ="ReportName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects."
                        "Type)=-32764)); "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =540
                            Width =975
                            Height =240
                            Name ="Label3"
                            Caption ="ReportName"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =180
                    Width =2475
                    ColumnWidth =3210
                    Name ="OrderByText"
                    ControlSource ="OrderByText"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =180
                            Width =1005
                            Height =240
                            Name ="Label5"
                            Caption ="OrderByText"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1920
                    Top =1200
                    TabIndex =1
                    Name ="Ascending"
                    ControlSource ="Ascending"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1200
                            Width =810
                            Height =240
                            Name ="Label7"
                            Caption ="Ascending"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
