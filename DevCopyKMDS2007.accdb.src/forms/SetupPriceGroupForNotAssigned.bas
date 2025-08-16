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
    Width =4380
    DatasheetFontHeight =10
    ItemSuffix =2
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
            Height =2145
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =60
                    Top =600
                    Width =3990
                    Height =480
                    FontSize =9
                    Name ="cmdStep1_SetupCategories"
                    Caption ="Step 1. Set Price Group IDENTITY_INSERT Off"
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
                            Caption ="Price Group"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =60
                    Top =1080
                    Width =4035
                    Height =480
                    FontSize =9
                    TabIndex =1
                    Name ="cmdStep2_SetupCategories"
                    Caption ="Step 2. Add Price Group 0 the set ID_INS ON "
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1740
                    Width =4260
                    Height =240
                    Name ="Label1"
                    Caption ="sqlPassThrough \"SET IDENTITY_INSERT SubCategory ON\""
                End
            End
        End
    End
End
CodeBehindForm
' See "SetupPriceGroupForNotAssigned.cls"
