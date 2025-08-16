Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =4215
    Top =2010
    Right =11415
    Bottom =5895
    RecSrcDt = Begin
        0x0dee43e59820e340
    End
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
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
        Begin Section
            Height =3900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =743
                    Top =60
                    Width =5415
                    Height =630
                    FontSize =24
                    Name ="Label0"
                    Caption =" Re - Assign an Item's ID"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =7200
                    Left =1680
                    Top =1380
                    Width =1500
                    Name ="Combo1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description, Item.Onhand, Item.Price FROM Item WHERE (((Ite"
                        "m.ID)<1000)) OR (((Item.ID)>99999)) ORDER BY Item.ID; "
                    ColumnWidths ="1440;2520;1440;600"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1200
                            Width =1470
                            Height =435
                            Name ="Item ID_Label"
                            Caption ="1. First choose the item to change."
                            EventProcPrefix ="Item_ID_Label"
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =7200
                    Left =3600
                    Top =1380
                    Width =3480
                    TabIndex =1
                    BoundColumn =1
                    Name ="cmbItemByDescription"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Item.ID, Item.Onhand, Item.Price FROM Item WHERE (((Ite"
                        "m.ID)<1000)) OR (((Item.ID)>99999)) ORDER BY Item.Description; "
                    ColumnWidths ="2880;1080;1440;598"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =2040
                    TabIndex =2
                    Name ="txtNewId"
                    Format ="General Number"
                    ValidationRule ="<1000 Or >99999"
                    ValidationText ="Item must be less than 1000 or greater than 99999"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =1860
                            Width =1470
                            Height =420
                            Name ="Label6"
                            Caption ="2. Enter the New Item ID."
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1680
                    Top =2640
                    Width =1170
                    Height =405
                    TabIndex =3
                    Name ="cmdChange"
                    Caption ="Change Item"
                    OnClick ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =2640
                    Width =1470
                    Height =420
                    Name ="Label8"
                    Caption ="3. Enter change Item."
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1830
                    Top =600
                    Width =3255
                    Height =390
                    FontSize =14
                    Name ="Label10"
                    Caption ="Related items only!"
                End
            End
        End
    End
End
CodeBehindForm
' See "ReAssignItemID.cls"
