VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo ExitHandler
    If Intersect(target, Me.Range("A:O")) Is Nothing Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop
    Application.ScreenUpdating = False

    Me.Unprotect
    
    Dim interCol As String
    Dim TargetLoc() As String
    Dim startRow As Integer
    Dim endRow As Long
    Dim lastRow As Long
    
    lastRow = Me.Cells(Rows.count, "A").End(xlUp).row
    
    If target.Cells.CountLarge = 1 Then
    
        CheckFacilities target
        
        ' check for special XML characters
         Application.Run "checkXMLchars", target
         
    ElseIf target.Rows.CountLarge = 1 Then
    
        
        If Intersect(Me.Range("N:N"), target) Is Nothing Then
            interCol = "A"
        ElseIf Intersect(Me.Range("O:O"), target) Is Nothing And _
            Intersect(Me.Range("P:P"), target) Is Nothing And _
            Intersect(Me.Range("Q:Q"), target) Is Nothing And _
            Intersect(Me.Range("R:R"), target) Is Nothing And _
            Intersect(Me.Range("S:S"), target) Is Nothing And _
            Intersect(Me.Range("T:T"), target) Is Nothing And _
            Intersect(Me.Range("U:U"), target) Is Nothing And _
            Intersect(Me.Range("V:V"), target) Is Nothing And _
            Intersect(Me.Range("W:W"), target) Is Nothing And _
            Intersect(Me.Range("X:X"), target) Is Nothing And _
            Intersect(Me.Range("Y:Y"), target) Is Nothing And _
            Intersect(Me.Range("Z:Z"), target) Is Nothing And _
            Intersect(Me.Range("AA:AA"), target) Is Nothing And _
            Intersect(Me.Range("AB:AB"), target) Is Nothing And _
            Intersect(Me.Range("AC:AC"), target) Is Nothing And _
            Intersect(Me.Range("AD:AD"), target) Is Nothing Then

            interCol = "N"
            
        Else
            interCol = "A"
        End If
        
        CheckFacilities Me.Range(interCol & target.row)
        
    ElseIf target.Rows.CountLarge < 1000 Then
    
    
        If endRow > lastRow Then Exit Sub
        
        TargetLoc = Split(target.Address(False, False), ":")
        
        startRow = Me.Range("A" & TargetLoc(0)).row
        endRow = Me.Range("A" & TargetLoc(1)).row
        
        
        If Intersect(Me.Range("N:N"), target) Is Nothing Then
            interCol = "A"
        ElseIf Intersect(Me.Range("O:O"), target) Is Nothing And _
            Intersect(Me.Range("P:P"), target) Is Nothing And _
            Intersect(Me.Range("Q:Q"), target) Is Nothing And _
            Intersect(Me.Range("R:R"), target) Is Nothing And _
            Intersect(Me.Range("S:S"), target) Is Nothing And _
            Intersect(Me.Range("T:T"), target) Is Nothing And _
            Intersect(Me.Range("U:U"), target) Is Nothing And _
            Intersect(Me.Range("V:V"), target) Is Nothing And _
            Intersect(Me.Range("W:W"), target) Is Nothing And _
            Intersect(Me.Range("X:X"), target) Is Nothing And _
            Intersect(Me.Range("Y:Y"), target) Is Nothing And _
            Intersect(Me.Range("Z:Z"), target) Is Nothing And _
            Intersect(Me.Range("AA:AA"), target) Is Nothing And _
            Intersect(Me.Range("AB:AB"), target) Is Nothing And _
            Intersect(Me.Range("AC:AC"), target) Is Nothing And _
            Intersect(Me.Range("AD:AD"), target) Is Nothing Then

            interCol = "N"
            
        Else
            interCol = "A"
        End If
        
        For row = startRow To endRow
        
            CheckFacilities Me.Range(interCol & row)
        
        Next row
        
     Else
     
        TargetLoc = Split(target.Address(False, False), ":")
        startRow = Me.Range("A" & TargetLoc(0)).row
        endRow = Me.Range("A" & TargetLoc(1)).row
        
        If endRow > lastRow Then Exit Sub
        
        If Intersect(Me.Range("N:N"), target) Is Nothing Then
            interCol = "B"
        ElseIf Intersect(Me.Range("O:O"), target) Is Nothing And _
            Intersect(Me.Range("P:P"), target) Is Nothing And _
            Intersect(Me.Range("Q:Q"), target) Is Nothing And _
            Intersect(Me.Range("R:R"), target) Is Nothing And _
            Intersect(Me.Range("S:S"), target) Is Nothing And _
            Intersect(Me.Range("T:T"), target) Is Nothing And _
            Intersect(Me.Range("U:U"), target) Is Nothing And _
            Intersect(Me.Range("V:V"), target) Is Nothing And _
            Intersect(Me.Range("W:W"), target) Is Nothing And _
            Intersect(Me.Range("X:X"), target) Is Nothing And _
            Intersect(Me.Range("Y:Y"), target) Is Nothing And _
            Intersect(Me.Range("Z:Z"), target) Is Nothing And _
            Intersect(Me.Range("AA:AA"), target) Is Nothing And _
            Intersect(Me.Range("AB:AB"), target) Is Nothing And _
            Intersect(Me.Range("AC:AC"), target) Is Nothing And _
            Intersect(Me.Range("AD:AD"), target) Is Nothing Then

            interCol = "N"
            
        Else
            interCol = "B"
        End If
        
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
        
            CheckFacilities Me.Range(interCol & row)
            
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
    Me.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Me.ScrollArea = ""
End Sub

Private Sub Worksheet_PivotTableBeforeCommitChanges(ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean)

End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If Intersect(target, Me.Range("AE:AE")) Is Nothing Then Exit Sub
    If target.Cells.CountLarge > 1 Then Exit Sub
    If ActiveCell.row < 4 Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop

    On Error GoTo ExitHandler

    AttForm.Show
    'ActiveCell.Offset(0, -1).Select

ExitHandler:

    
    Me.ScrollArea = "A:AG"
    Application.EnableEvents = True
    
End Sub
