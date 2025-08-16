Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =-2820
    RecSrcDt = Begin
        0xf5502a2bcee7e340
    End
    RecordSource ="qryKitsWithTotalsAndItems"
    Caption ="qryKitsWithTotalsAndItems"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =204
            FontSize =9
            FontWeight =700
            BorderColor =5066944
            ForeColor =11830108
            FontName ="Calibri"
            GridlineColor =-2147483609
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
            BackColor =-2147483617
            BorderColor =-2147483617
            GridlineColor =-2147483609
        End
        Begin Line
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin Image
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BackColor =-2147483617
            BorderColor =-2147483617
            GridlineColor =-2147483609
        End
        Begin CommandButton
            TextFontCharSet =204
            Width =1701
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483609
            BorderLineStyle =0
        End
        Begin OptionButton
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            GridlineColor =-2147483609
        End
        Begin OptionGroup
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5066944
        End
        Begin BoundObjectFrame
            SizeMode =3
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =204
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            AsianLineBreak =1
            GridlineColor =-2147483609
            ShowDatePicker =0
        End
        Begin ListBox
            TextFontCharSet =204
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =204
            BorderLineStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            BorderColor =5066944
            FontName ="Calibri"
            GridlineColor =-2147483609
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin UnboundObjectFrame
            BackStyle =0
            Width =4536
            Height =2835
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =5066944
            GridlineColor =-2147483609
        End
        Begin ToggleButton
            TextFontCharSet =204
            Width =283
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =-2147483609
            FontName ="Calibri"
            GridlineColor =-2147483609
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            BackStyle =0
            Width =5103
            Height =3402
            FontSize =9
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =1701
            Height =1701
            BorderColor =5066944
            GridlineColor =-2147483609
            LabelX =-1701
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="KitItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =582
            Name ="ReportHeader"
            AutoHeight =1
            Begin
                Begin Label
                    TextFontFamily =18
                    Left =57
                    Top =57
                    Width =5265
                    Height =525
                    FontSize =20
                    ForeColor =-2147483617
                    Name ="Label27"
                    Caption ="qryKitsWithTotalsAndItems"
                    FontName ="Cambria"
                    GridlineColor =0
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =5322
                    LayoutCachedHeight =582
                End
            End
        End
        Begin PageHeader
            Height =737
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            Name ="GroupHeader2"
            AutoHeight =255
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Top =30
                    Width =1635
                    Height =345
                    FontSize =12
                    FontWeight =700
                    ForeColor =11830108
                    Name ="qryKitsOnFile.ItemID"
                    ControlSource ="KitItemID"
                    EventProcPrefix ="qryKitsOnFile_ItemID"
                    GridlineColor =0

                    LayoutCachedTop =30
                    LayoutCachedWidth =1635
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2077
                    Width =3075
                    Height =345
                    FontSize =12
                    TabIndex =1
                    Name ="qryKitsOnFile.Description"
                    ControlSource ="KitDescription"
                    EventProcPrefix ="qryKitsOnFile_Description"

                    LayoutCachedLeft =2077
                    LayoutCachedWidth =5152
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6066
                    Width =1500
                    Height =345
                    FontSize =12
                    TabIndex =2
                    Name ="qryKitsOnFile.Price"
                    ControlSource ="KitPrice"
                    EventProcPrefix ="qryKitsOnFile_Price"

                    LayoutCachedLeft =6066
                    LayoutCachedWidth =7566
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7707
                    Width =1311
                    Height =345
                    FontSize =12
                    TabIndex =3
                    Name ="RetailPrice"
                    ControlSource ="RetailPrice"

                    LayoutCachedLeft =7707
                    LayoutCachedWidth =9018
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9297
                    Width =1311
                    Height =345
                    FontSize =12
                    TabIndex =4
                    Name ="Discount"
                    ControlSource ="KitDiscount"

                    LayoutCachedLeft =9297
                    LayoutCachedWidth =10608
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =10549
                    Width =966
                    Height =345
                    ColumnWidth =2205
                    FontSize =12
                    TabIndex =5
                    Name ="KitDiscPerct"
                    ControlSource ="KitDiscPerct"
                    Format ="Percent"

                    LayoutCachedLeft =10549
                    LayoutCachedWidth =11515
                    LayoutCachedHeight =345
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =348
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =3905
                    Width =2610
                    Height =270
                    Name ="PartDescription"
                    ControlSource ="PartDescription"

                    LayoutCachedLeft =3905
                    LayoutCachedWidth =6515
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2607
                    Width =1245
                    Height =270
                    TabIndex =1
                    Name ="PartID"
                    ControlSource ="PartID"

                    LayoutCachedLeft =2607
                    LayoutCachedWidth =3852
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7653
                    Width =1365
                    Height =270
                    TabIndex =2
                    Name ="PartPrice"
                    ControlSource ="PartPrice"

                    LayoutCachedLeft =7653
                    LayoutCachedWidth =9018
                    LayoutCachedHeight =270
                End
            End
        End
        Begin PageFooter
            Height =453
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
