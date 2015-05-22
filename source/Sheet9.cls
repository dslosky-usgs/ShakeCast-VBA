VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Me.Unprotect
    Me.Rows("1:36").Select
    
    ActiveWindow.Zoom = True
    
    Application.Run "protectWorkbook"
    
    Me.Range("Z200").Activate
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)

End Sub











