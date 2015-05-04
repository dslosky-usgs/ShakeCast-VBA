VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    'If Intersect(Target, Me.Range("A:A")) Is Nothing Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop
    Application.ScreenUpdating = False
    
    Me.Unprotect
    
On Error GoTo ExitHandler

If target.Cells.count = 1 Then

    CheckUsers target.row

    ' check for special XML characters
     Application.Run "checkXMLchars", target
     
Else

    Dim TargetLoc() As String
    Dim startRow As Integer
    Dim endRow As Integer
    
    TargetLoc = Split(target.Address(False, False), ":")
    startRow = Me.Range(TargetLoc(0)).row
    endRow = Me.Range(TargetLoc(1)).row
    
    For row = startRow To endRow
    
        CheckUsers row
    
    Next row


End If




ExitHandler:
    Me.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Me.ScrollArea = ""
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    On Error GoTo ExitHandler
    If Intersect(target, Me.Range("G:G")) Is Nothing Then GoTo checkPhone
        If target.row < 4 Then Exit Sub
        If target.count > 1 Then Exit Sub
        
        Application.EnableEvents = False 'to prevent endless loop
    
        
    
        Application.Run "makeGroupChecklist"
        GroupCheckBox.Show
        'ActiveCell.Offset(0, 1).Select
        
checkPhone:
    If Intersect(target, Me.Range("F:F")) Is Nothing Then GoTo ExitHandler
    
    PhoneForm.Show
    
    
ExitHandler:
    Worksheets("User XML").ScrollArea = ""
    Application.EnableEvents = True

End Sub


