Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command2_Click()

CreateWinSCPtoKMDSFTPScriptFile


End Sub

Private Sub Command3_Click()
    SetOptions False
    WaitForExecCmd "c:\npos\WinSCPKS.BAT"
    DoCmd.RunSQL "Delete * from KMDSKS"
    DoCmd.TransferText acImportFixed, "KMDSSWImportSpecification", "KMDSKS", "K:\Mail\In\KMDSSW.TXT", False
    SetOptions True
    DoCmd.Requery

End Sub
