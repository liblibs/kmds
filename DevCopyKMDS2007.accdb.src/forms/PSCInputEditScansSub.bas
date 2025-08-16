Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14160
    DatasheetFontHeight =10
    ItemSuffix =27
    Filter ="((PSCInput.ItemID=84902))"
    RecSrcDt = Begin
        0x701725d9eedde240
    End
    RecordSource ="PSCInput"
    Caption ="PSCInput"
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
            Height =300
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =900
                    ColumnWidth =1575
                    ColumnOrder =0
                    Name ="Code"
                    ControlSource ="Code"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =885
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Quantity"
                    ControlSource ="Quantity"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8580
                    Top =60
                    Width =900
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =5
                    TabIndex =2
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9600
                    Top =60
                    Width =1140
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =6
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"
                    Format ="$#,##0.00;($#,##0.00)"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10860
                    Top =60
                    Width =840
                    Height =255
                    ColumnWidth =870
                    ColumnOrder =7
                    TabIndex =4
                    Name ="Cost"
                    ControlSource ="Cost"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =2040
                    Top =60
                    Width =1020
                    ColumnOrder =2
                    TabIndex =5
                    Name ="Item Code"
                    ControlSource ="ItemID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Item].[ID], [Item].[Description], [Item].[Price] FROM Item ORDER BY [ID]"
                        "; "
                    ColumnWidths ="1440;1440;1440"
                    BeforeUpdate ="[Event Procedure]"
                    EventProcPrefix ="Item_Code"

                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =3180
                    Top =60
                    Width =3240
                    ColumnWidth =3585
                    ColumnOrder =3
                    TabIndex =6
                    Name ="Item Description"
                    ControlSource ="Description"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Item].[Description], [Item].[ID], [Item].[Price] FROM Item ORDER BY [Des"
                        "cription]; "
                    ColumnWidths ="1440;1440;1440"
                    BeforeUpdate ="[Event Procedure]"
                    EventProcPrefix ="Item_Description"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =6480
                    Top =60
                    Width =1980
                    ColumnOrder =4
                    TabIndex =7
                    BoundColumn =1
                    Name ="Sales Dpt"
                    ControlSource ="SalesDepartmentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [ID] & \" - \" & [Description] AS [Sales Dpt], SalesDepartment.ID FROM Sa"
                        "lesDepartment ORDER BY [ID] & \" - \" & [Description]; "
                    ColumnWidths ="1440;1440;1440"
                    EventProcPrefix ="Sales_Dpt"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11820
                    Top =120
                    TabIndex =8
                    Name ="Locked"
                    ControlSource ="Locked"

                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12180
                    Top =60
                    Width =1980
                    TabIndex =9
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

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
                    Left =1620
                    Width =900
                    Height =255
                    Name ="Text20"
                    ControlSource ="=Sum([Quantity])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =1560
                            Height =255
                            Name ="Label21"
                            Caption ="Quantity"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Width =2820
                    Height =255
                    TabIndex =1
                    Name ="Text22"
                    ControlSource ="=Sum(([Quantity])*[Price])"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Width =1560
                            Height =255
                            Name ="Label23"
                            Caption ="Quantity"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PSCInputEditScansSub.cls"
