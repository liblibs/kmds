Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6240
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =3195
    Top =7065
    Right =11430
    Bottom =11790
    OrderBy ="LineID"
    RecSrcDt = Begin
        0xc9f9ff9f8c9fe240
    End
    RecordSource ="SELECT ClubOrderGuideTemp.ClubID, ClubOrderGuideTemp.LineID, ClubOrderGuideTemp."
        "ItemId, ClubOrderGuideTemp.Par, ClubOrderGuideTemp.QuotedPrice, ClubOrderGuideTe"
        "mp.LiquorItem FROM ClubOrderGuideTemp WHERE (((ClubOrderGuideTemp.LiquorItem)=Tr"
        "ue)) ORDER BY ClubOrderGuideTemp.LineID, ClubOrderGuideTemp.ItemId; "
    Caption ="ClubOrderGuideTemp"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    OnKeyDown ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    FilterOnLoad =0
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Name ="Club"
                    ControlSource ="ClubID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =60
                            Width =390
                            Height =240
                            Name ="Label1"
                            Caption ="Club"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =2460
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2640
                    Top =120
                    ColumnWidth =720
                    ColumnOrder =0
                    Name ="LineID"
                    ControlSource ="LineID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =120
                            Width =585
                            Height =240
                            Name ="Label9"
                            Caption ="Line #"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =25
                    ListWidth =5760
                    Left =2640
                    Top =540
                    ColumnWidth =780
                    ColumnOrder =1
                    TabIndex =1
                    Name ="ItemId"
                    ControlSource ="ItemId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description, Item.Price, Size.Description AS SizeD, Item.Sa"
                        "lesDepartmentID FROM Item LEFT JOIN [Size] ON Item.SizeID=Size.ID WHERE (((Item."
                        "SalesDepartmentID)=1)) ORDER BY Item.ID; "
                    ColumnWidths ="1080;2880;720;720"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =540
                            Width =615
                            Height =240
                            Name ="Label10"
                            Caption ="Code #"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =25
                    ListWidth =5760
                    Left =2700
                    Top =960
                    Width =2040
                    ColumnWidth =1170
                    ColumnOrder =2
                    TabIndex =2
                    BoundColumn =1
                    Name ="ItemDesc"
                    ControlSource ="ItemId"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Item.ID, Price, Size.Description AS SizeD, Item.SalesDe"
                        "partmentID FROM Item LEFT JOIN [Size] ON Item.SizeID=Size.ID WHERE (((Item.Sales"
                        "DepartmentID)=1)) ORDER BY Item.Description; "
                    ColumnWidths ="2880;1080;720;720"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =960
                            Width =870
                            Height =240
                            Name ="Label12"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2700
                    Top =1380
                    ColumnWidth =540
                    ColumnOrder =3
                    TabIndex =3
                    Name ="Par"
                    ControlSource ="Par"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1380
                            Width =375
                            Height =240
                            Name ="Label13"
                            Caption ="Par:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2700
                    Top =1800
                    ColumnWidth =1320
                    TabIndex =4
                    Name ="QuotedPrice"
                    ControlSource ="QuotedPrice"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1800
                            Width =1020
                            Height =240
                            Name ="Label14"
                            Caption ="QuotedPrice:"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2520
                    Top =2220
                    ColumnWidth =780
                    TabIndex =5
                    Name ="LiquorItem"
                    ControlSource ="LiquorItem"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1620
                            Top =2220
                            Width =900
                            Height =240
                            Name ="Label15"
                            Caption ="LiquorItem:"
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
' See "ClubOrderGuideTempSub.cls"
