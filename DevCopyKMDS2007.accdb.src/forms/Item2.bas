Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12240
    DatasheetFontHeight =10
    ItemSuffix =28
    Top =600
    Bottom =600
    RecSrcDt = Begin
        0x10edc6b4d3aae240
    End
    RecordSource ="Item"
    Caption ="Item2"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =790
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =25
                    ListWidth =5760
                    Left =1140
                    Top =180
                    Width =1680
                    Name ="cmbItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.ID, Item.Description, Item.Onhand, Item.Price FROM Item ORDER BY Ite"
                        "m.ID; "
                    ColumnWidths ="1440;1440;1440;1440"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =180
                            Width =630
                            Height =245
                            Name ="Item ID_Label"
                            Caption ="Item ID"
                            EventProcPrefix ="Item_ID_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =25
                    ListWidth =7080
                    Left =4740
                    Top =180
                    Width =2880
                    TabIndex =1
                    BoundColumn =1
                    Name ="cmbDesc"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Item.Description, Item.ID, Item.Onhand, Item.Price FROM Item ORDER BY Ite"
                        "m.Description; "
                    ColumnWidths ="2760;1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =180
                            Width =1500
                            Height =245
                            Name ="Description Lookup_Label"
                            Caption ="Description Lookup"
                            EventProcPrefix ="Description_Lookup_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =25
                    ListWidth =7080
                    Left =9360
                    Top =180
                    Width =2100
                    TabIndex =2
                    BoundColumn =2
                    Name ="cmbUPC"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT UPCItem.ID AS UPCCode, Item.Description, Item.ID, Item.Onhand, Item.Price"
                        " FROM Item INNER JOIN UPCItem ON Item.ID=UPCItem.ItemID ORDER BY UPCItem.ID; "
                    ColumnWidths ="2759;1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7740
                            Top =180
                            Width =1500
                            Height =245
                            Name ="Label27"
                            Caption ="UPC Lookup"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =4920
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =120
                    Width =900
                    Height =255
                    ColumnWidth =900
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =120
                            Width =1560
                            Height =255
                            Name ="ID_Label"
                            Caption ="ID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Top =120
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2760
                            Top =120
                            Width =1560
                            Height =255
                            Name ="Description_Label"
                            Caption ="Description"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1680
                    Top =840
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =840
                            Width =1560
                            Height =255
                            Name ="Onhand_Label"
                            Caption ="Onhand"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1680
                    Top =1200
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =3
                    Name ="OrderAt"
                    ControlSource ="OrderAt"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1200
                            Width =1560
                            Height =255
                            Name ="OrderAt_Label"
                            Caption ="OrderAt"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1680
                    Top =1560
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =4
                    Name ="OrderTo"
                    ControlSource ="OrderTo"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1560
                            Width =1560
                            Height =255
                            Name ="OrderTo_Label"
                            Caption ="OrderTo"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3255
                    Top =840
                    Width =780
                    Height =255
                    ColumnWidth =2310
                    TabIndex =5
                    Name ="Price"
                    ControlSource ="Price"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2760
                            Top =840
                            Width =435
                            Height =240
                            Name ="Price_Label"
                            Caption ="Price"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3255
                    Top =1200
                    Width =780
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Cost"
                    ControlSource ="Cost"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2805
                            Top =1200
                            Width =390
                            Height =240
                            Name ="Cost_Label"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5445
                    Top =840
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =7
                    Name ="UnitsPerCase"
                    ControlSource ="UnitsPerCase"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4350
                            Top =840
                            Width =1035
                            Height =240
                            Name ="UnitsPerCase_Label"
                            Caption ="UnitsPerCase"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5445
                    Top =1200
                    Width =900
                    Height =255
                    ColumnWidth =900
                    TabIndex =8
                    Name ="UnitsPerOrder"
                    ControlSource ="UnitsPerOrder"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4320
                            Top =1200
                            Width =1065
                            Height =240
                            Name ="UnitsPerOrder_Label"
                            Caption ="UnitsPerOrder"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =315
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =60
                    Width =1920
                    Height =255
                    ColumnWidth =1035
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =1560
                            Height =255
                            Name ="SetupDate_Label"
                            Caption ="SetupDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Width =2340
                    Height =255
                    ColumnWidth =1035
                    TabIndex =1
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3960
                            Top =60
                            Width =1560
                            Height =255
                            Name ="ChangeDate_Label"
                            Caption ="ChangeDate"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "Item2.cls"
