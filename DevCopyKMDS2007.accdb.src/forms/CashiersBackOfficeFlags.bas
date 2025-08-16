Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10320
    DatasheetFontHeight =10
    ItemSuffix =22
    Left =525
    Top =1230
    Right =525
    Bottom =1230
    RecSrcDt = Begin
        0xab20b705f79de240
    End
    RecordSource ="Cashier"
    Caption ="Cashiers BackOffice Flags"
    OnOpen ="[Event Procedure]"
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
            Height =480
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =420
                    Height =420
                    FontWeight =700
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =540
                    Top =60
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="Name_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1620
                    Top =60
                    Width =900
                    Height =420
                    FontWeight =700
                    Name ="AllowChangePrice_Label"
                    Caption ="Change Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2580
                    Top =60
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="AllowChangeOnhand_Label"
                    Caption ="Change Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3660
                    Top =60
                    Width =840
                    Height =420
                    FontWeight =700
                    Name ="AllowChangeCost_Label"
                    Caption ="Change Cost"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4560
                    Top =60
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="AllowChangeOrderAt_Label"
                    Caption ="Change Order At"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5640
                    Top =60
                    Width =1020
                    Height =420
                    FontWeight =700
                    Name ="AllowChangeOrderTo_Label"
                    Caption ="Change Order To"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6720
                    Top =60
                    Width =1140
                    Height =420
                    FontWeight =700
                    Name ="AllowChangeDescription_Label"
                    Caption ="Change Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7980
                    Top =60
                    Width =1140
                    Height =420
                    FontWeight =700
                    Name ="Label16"
                    Caption ="View Hourly Wages"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7980
                    Top =60
                    Width =1140
                    Height =420
                    FontWeight =700
                    Name ="Label18"
                    Caption ="View Hourly Wages"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9180
                    Top =60
                    Width =1140
                    Height =420
                    FontWeight =700
                    Name ="Label20"
                    Caption ="Delete Documents"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =315
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =600
                    BackColor =8421631
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =1020
                    Height =255
                    ColumnWidth =1560
                    TabIndex =1
                    BackColor =8421631
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2040
                    Top =60
                    TabIndex =2
                    Name ="AllowChangePrice"
                    ControlSource ="AllowChangePrice"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3000
                    Top =60
                    TabIndex =3
                    Name ="AllowChangeOnhand"
                    ControlSource ="AllowChangeOnhand"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3960
                    Top =60
                    TabIndex =4
                    Name ="AllowChangeCost"
                    ControlSource ="AllowChangeCost"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4980
                    Top =60
                    TabIndex =5
                    Name ="AllowChangeOrderAt"
                    ControlSource ="AllowChangeOrderAt"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6060
                    Top =60
                    TabIndex =6
                    Name ="AllowChangeOrderTo"
                    ControlSource ="AllowChangeOrderTo"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7140
                    Top =60
                    TabIndex =7
                    Name ="AllowChangeDescription"
                    ControlSource ="AllowChangeDescription"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8400
                    Top =60
                    TabIndex =8
                    Name ="Check17"
                    ControlSource ="AllowHourlyWageView"

                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =8400
                    Top =60
                    TabIndex =9
                    Name ="Check19"
                    ControlSource ="AllowHourlyWageView"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9600
                    Top =60
                    TabIndex =10
                    Name ="Check21"
                    ControlSource ="AllowDeleteDocument"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "CashiersBackOfficeFlags.cls"
