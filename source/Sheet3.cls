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

Dim TargetLoc() As String
Dim startRow As Integer
Dim endRow As Integer

If target.Cells.count = 1 Then

    CheckUsers target.row

    ' check for special XML characters
     Application.Run "checkXMLchars", target
     
ElseIf target.Rows.count = 1 Then

    CheckUsers target.row

ElseIf target.Rows.count < 100 Then
    TargetLoc = Split(target.Address(False, False), ":")
    startRow = Me.Range("A" & TargetLoc(0)).row
    endRow = Me.Range("A" & TargetLoc(1)).row
    
    For row = startRow To endRow
    
        CheckUsers row
    
    Next row
    
Else
    
    TargetLoc = Split(target.Address(False, False), ":")
    startRow = Me.Range("A" & TargetLoc(0)).row
    endRow = Me.Range("A" & TargetLoc(1)).row
    
    ' set up progress bar
    Dim progressCount As Long
    Dim progressWhen As Long
    Dim pcntDone As Double
   
    progressCount = 0
    progressWhen = (endRow - startRow) * 0.01
    pcntDone = 0
    ProgressForm.ProcessName.Caption = "Updating Worksheet"
    ProgressForm.ProgressLabel.Width = 0
   
    ProgressForm.Show vbModeless
   
    For row = startRow To endRow
   
        CheckUsers row
       
        If progressCount > progressWhen Then
       
            pcntDone = ((row - startRow - 4) / (endRow - startRow - 3))
           
            ProgressForm.ProgressLabel.Width = pcntDone * ProgressForm.MaxWidth.Caption
            ProgressForm.ProgressFrame.Caption = Round(pcntDone * 100, 0) & "%"
           
            DoEvents
            progressCount = 0
       
        End If
       
        progressCount = progressCount + 1
   
    Next row
   
   
    Unload ProgressForm

End If






ExitHandler:
    Unload ProgressForm
    Me.Cells.Locked = False
    Me.Rows(1).Locked = True
    Me.Rows(2).Locked = True
    Me.Rows(3).Locked = True
    Me.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True
    Me.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = False
    Application.EnableEvents = True
    Me.ScrollArea = ""
    Application.Run (protectWorkbook)
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    ' On Error GoTo ExitHandler
    If Intersect(target, Me.Range("G:G")) Is Nothing Then GoTo checkPhone
        If target.row < 4 Then Exit Sub
        If target.count > 1 Then Exit Sub
        
        Application.EnableEvents = False 'to prevent endless loop
    
        GroupCheckBox.Show
        'ActiveCell.Offset(0, 1).Select
        
checkPhone:
    If Intersect(target, Me.Range("F:F")) Is Nothing Then GoTo UserDropdown
    
    PhoneForm.Show
    
UserDropdown:
    If Intersect(target, Me.Range("B:B")) Is Nothing Then GoTo ExitHandler
    
    Me.Unprotect
    
    ' create a string array with the user types
    Dim UserTypes(0 To 1) As String
    UserTypes(0) = "USER"
    UserTypes(1) = "ADMIN"
    
    ' create user/admin drop down
    With Me.Range("B" & target.row).Validation
        .delete
        .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Join(UserTypes, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "User Type"
        .ErrorTitle = ""
        .InputMessage = "Please select a user type type from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
ExitHandler:
    Worksheets("User XML").ScrollArea = ""
    Application.EnableEvents = True
    Application.Run "ProtectWorkbook"

End Sub














