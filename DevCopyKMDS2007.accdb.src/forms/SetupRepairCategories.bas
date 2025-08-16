Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
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
    Width =5160
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =1155
    Top =3645
    Right =6810
    Bottom =6540
    RecSrcDt = Begin
        0x5469632d1843e340
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
        Begin Section
            Height =2100
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =60
                    Top =600
                    Width =4920
                    Height =480
                    FontSize =9
                    Name ="cmdStep1_SetupCategories"
                    Caption ="Step 1. sp_Setup_ID_0_SubCategory_Vendor_PriceGroup"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =1560
                            Height =420
                            FontSize =12
                            Name ="Label0"
                            Caption ="setup 0 id's"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =60
                    Top =1560
                    Width =4980
                    Height =480
                    FontSize =9
                    TabIndex =1
                    Name ="cmdStep3_SetupOLCCCategories"
                    Caption ="Step3. Add  Categories as per OLCC (Condensed Version)"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =2040
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =60
                    Top =1080
                    Width =4860
                    Height =480
                    FontSize =9
                    TabIndex =2
                    Name ="cmdStep2_SetupCategories"
                    Caption ="Step 2. Update Item Subcategory Vendor Price Group 0 "
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "SetupRepairCategories.cls"
