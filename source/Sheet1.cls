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
    If Intersect(target, Me.Range("A:M")) Is Nothing Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop
    Application.ScreenUpdating = Falserr
    Me.Unprotect
    
    Dim interCol As String
    Dim TargetLoc() As String
    Dim startRow As Integer
    Dim endRow As Long
    Dim lastRow As Long
    
    lastRow = Me.Cells(Rows.count, "A").End(xlUp).row
    If target.Cells.CountLarge = 1 Then
        
        If Intersect(Me.Range("facility_mbt"), target) Is Nothing Then
            CheckFacilities Me.Range("A" & target.row)
        Else
            CheckFacilities target
        End If
        
        ' check for special XML characters
         Application.Run "checkXMLchars", target
         
    ElseIf target.Rows.CountLarge = 1 Then
        If Intersect(Me.Range("facility_mbt"), target) Is Nothing Then
            interCol = "A"
        ElseIf Intersect(Me.Range("facility_fragility"), target) Is Nothing Then
            interCol = "M"
        Else
            interCol = "A"
        End If
        
        CheckFacilities Me.Range(interCol & target.row)
        
    ElseIf target.Rows.CountLarge < 100 Then
        If endRow > lastRow Then Exit Sub
        
        TargetLoc = Split(target.Address(False, False), ":")
        startRow = Me.Range("A" & TargetLoc(0)).row
        endRow = Me.Range("A" & TargetLoc(1)).row
        
        If Intersect(Me.Range("facility_mbt"), target) Is Nothing Then
            interCol = "A"
        ElseIf Intersect(Me.Range("facility_fragility"), target) Is Nothing Then
            interCol = "M"
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
        
        If endRow > lastRow Then GoTo ExitHandler
        
        If Intersect(Me.Range("N:N"), target) Is Nothing Then
            interCol = "B"
        ElseIf Intersect(Me.Range("facility_fragility"), target) Is Nothing Then
            interCol = "M"
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
     
     End If
     

     
ExitHandler:
    Unload ProgressForm
    Me.Cells.Locked = False
    Me.Rows(1).Locked = True
    Me.Rows(2).Locked = True
    Me.Rows(3).Locked = True
    Me.Protect AllowFormattingCells:=True, AllowDeletingRows:=True, AllowInsertingRows:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Me.ScrollArea = ""
End Sub

Private Sub Worksheet_PivotTableBeforeCommitChanges(ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean)

End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If target.Cells.CountLarge > 1 Then Exit Sub
    If target.row < 4 Then Exit Sub
    Application.EnableEvents = False 'to prevent endless loop
    
    Me.Unprotect
    
    On Error GoTo ExitHandler

    Dim rowNum As Integer
    Dim colNum As Integer
    
    rowNum = target.row
    colNum = target.Column
    
    Set LookUpSheet = Worksheets("ShakeCast Ref Lookup Values")

    If colNum <> 30 Then GoTo FacDrop
    
        AttForm.Show

FacDrop:
    If colNum <> 2 Then GoTo ModDrop
    
        Set FacType = Me.Range("B" & rowNum)
        
        Dim FacTypes() As String
        
        lastFac = LookUpSheet.Cells(Rows.count, "C").End(xlUp).row
        Set FacTypeCells = Worksheets("ShakeCast Ref Lookup Values").Range("C1:C" & lastFac)

        With FacType.Validation
            .delete
            .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="='" & LookUpSheet.Name & "'!" & FacTypeCells.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Facility Type"
            .ErrorTitle = ""
            .InputMessage = "Please select a facility type from the drop-down list"
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    
ModDrop:
    If colNum <> 13 Then GoTo ExitHandler
    
        Set ModType = Me.Range("M" & rowNum)
        
        Dim ModTypes() As String
    
        Set hazSheet = Worksheets("Hazus Facility Model Data")
        Dim lastHaz As Long
        
        lastHaz = hazSheet.Cells(Rows.count, "A").End(xlUp).row
        
        Set ModTypeCells = hazSheet.Range("A2:A" & lastHaz)
        
        With ModType.Validation
            .delete
            .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="='" & hazSheet.Name & "'!" & ModTypeCells.Address
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "HAZUS Model Type"
            .ErrorTitle = ""
            .InputMessage = "Please select a model type from the drop-down list"
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    
ExitHandler:
    Me.ScrollArea = ""
    Application.EnableEvents = True
    Application.Run "ProtectWorkbook"
    
End Sub






















