Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =27
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6780
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =-390
    Top =4365
    Right =9855
    Bottom =11145
    DatasheetBackColor =65535
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0xfe3e318f8de3e440
    End
    RecordSource ="qryClubOrderGuideWeb"
    Caption ="ClubOrderGuides"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    DatasheetBackColor12 =65535
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
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
            Height =720
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =600
                    ColumnOrder =0
                    Name ="Order"
                    ControlSource ="OrderID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2280
                            Top =60
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="Order"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3660
                    Top =90
                    ColumnOrder =1
                    TabIndex =1
                    Name ="chkNonLiquor"
                    ControlSource ="NonLiquor"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3890
                            Top =60
                            Width =855
                            Height =240
                            Name ="Label22"
                            Caption ="NonLiquor"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3660
                    Top =480
                    ColumnOrder =2
                    TabIndex =2
                    Name ="chkMaintainParLevels"
                    ControlSource ="MaintainParLevels"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3890
                            Top =450
                            Width =855
                            Height =240
                            Name ="Label24"
                            Caption ="NonLiquor"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =3240
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =120
                    ColumnWidth =-2
                    Name ="LineID"
                    ControlSource ="LineID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =120
                            Width =525
                            Height =240
                            Name ="Label5"
                            Caption ="Line #"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =600
                    ColumnWidth =-2
                    TabIndex =1
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =600
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="Code #"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2220
                    Top =1080
                    Width =2760
                    ColumnWidth =-2
                    TabIndex =2
                    Name ="Combo16"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description FROM Item ORDER BY Item.ID; "
                    ColumnWidths ="0;2880"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =780
                            Top =1080
                            Width =930
                            Height =240
                            Name ="Label17"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Top =2160
                    ColumnWidth =780
                    TabIndex =3
                    Name ="QuotedPrice"
                    ControlSource ="QuotedPrice"

                    LayoutCachedLeft =3660
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =2400
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2220
                            Top =2160
                            Width =1020
                            Height =240
                            Name ="Label19"
                            Caption ="Quote:"
                            LayoutCachedLeft =2220
                            LayoutCachedTop =2160
                            LayoutCachedWidth =3240
                            LayoutCachedHeight =2400
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =1680
                    Height =255
                    ColumnWidth =-2
                    TabIndex =4
                    Name ="Par"
                    ControlSource ="Par"
                    Format ="Fixed"

                    LayoutCachedLeft =3300
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =1935
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1860
                            Top =1680
                            Width =375
                            Height =240
                            Name ="Label25"
                            Caption ="Par:"
                            LayoutCachedLeft =1860
                            LayoutCachedTop =1680
                            LayoutCachedWidth =2235
                            LayoutCachedHeight =1920
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2820
                    ColumnWidth =0
                    ColumnOrder =0
                    TabIndex =5
                    Name ="ClubID"
                    ControlSource ="ClubID"

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1080
                            Top =2820
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="ClubID"
                            LayoutCachedLeft =1080
                            LayoutCachedTop =2820
                            LayoutCachedWidth =1635
                            LayoutCachedHeight =3060
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
CodeBehindForm
' See "ClubOrderGuidesWeb.cls"
