VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo ExitHandler
    'If Intersect(Target, Me.Range("A:A")) Is Nothing Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop
    Application.ScreenUpdating = False
    
    Me.Activate
    Me.Unprotect


    CheckGroups

    
    If target.count = 1 Then
        ' check for special XML characters
        Application.Run "checkXMLchars", target
        
        ' skip over grey cells
        If Me.Range("A" & target.row).Value = Me.Range("A" & target.row - 1).Value Then
            If target.column = 1 Then
                Me.Range("D" & target.row).Activate
            ElseIf target.column = 4 And target.Value = "NEW_EVENT" Then
                Me.Range("F" & target.row).Activate
            ElseIf target.column = 5 And target.Value = "DAMAGE" Then
                Me.Range("G" & target.row).Activate
            End If
        End If
    End If
    
ExitHandler:
    Application.Run "protectWorkbook"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Me.ScrollArea = ""
End Sub
Private Sub Worksheet_SelectionChange(ByVal target As Range)
    On Error GoTo ExitHandler
    Application.EnableEvents = False
    If target.column <> 2 Or Me.Range("A2").Value <> "Advanced User" Then GoTo ExitHandler
    
    Application.Run "makeFacTypeChecklist"
    
ExitHandler:
    Application.EnableEvents = True
End Sub






