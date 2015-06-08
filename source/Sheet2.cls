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
        If Me.Range("A" & target.row).value = Me.Range("A" & target.row - 1).value Then
            If target.Column = 1 Then
                Me.Range("D" & target.row).Activate
            ElseIf target.Column = 4 And target.value = "NEW_EVENT" Then
                Me.Range("F" & target.row).Activate
            ElseIf target.Column = 5 And target.value = "DAMAGE" Then
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
    If target.row < 4 Then Exit Sub
    
    Application.EnableEvents = False
    
    Me.Unprotect
    
    '''''''''''''''''''''''''''''''''''''' setup variables '''''''''''''''''''''''''''''''''''''''''''
    Dim rowNum As Integer
    Dim colNum As Integer
    
    rowNum = target.row
    colNum = target.Column
    
    Set LookUpSheet = Worksheets("ShakeCast Ref Lookup Values")
    
    
    '''''''''''''''''''''''''''''''''''''' Make Drop Downs '''''''''''''''''''''''''''''''''''''''''''
    
    If target.Cells.count > 1 Then GoTo ExitHandler
    
    If target.Column <> 2 Or Me.Range("A2").value <> "Advanced User" Then GoTo FacilityDrop
    
        Application.Run "makeFacTypeChecklist"
    
' create facility drop down menu
FacilityDrop:
    If colNum <> 2 Then GoTo NotTypeDrop
    
            Set FacType = Me.Range("B" & (rowNum))
            
            lastFac = LookUpSheet.Cells(Rows.count, "C").End(xlUp).row
            Set FacTypeCells = LookUpSheet.Range("C1:C" & lastFac)
            
            With FacType.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
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
    
NotTypeDrop:
    If colNum <> 4 Then GoTo PrioDrop
        
        Set NotType = Me.Range("D" & (rowNum))
        ' create an array to hold the items for the Notification Type
        Dim NotTypes(0 To 1) As String
        NotTypes(0) = "NEW_EVENT"
        NotTypes(1) = "DAMAGE"
        
        With NotType.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=Join(NotTypes, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Notification Type"
            .ErrorTitle = ""
            .InputMessage = "Please select a notification type from the drop-down list"
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    
PrioDrop:
    If colNum <> 5 Then GoTo EventDrop
    
    Set InsPrio = Me.Range("E" & (rowNum))
    ' create an array for inspection priority
    Dim InsPrios(0 To 3) As String
    InsPrios(0) = "GREEN"
    InsPrios(1) = "YELLOW"
    InsPrios(2) = "ORANGE"
    InsPrios(3) = "RED"
    
    With InsPrio.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Join(InsPrios, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Event Type"
        .ErrorTitle = ""
        .InputMessage = "Please select an event type from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
EventDrop:
    If colNum <> 7 Then GoTo NotFormatDrop
    
    Set EvType = Me.Range("G" & (rowNum))
    
    ' create an array for event type
    Dim EvTypeStr() As String
    If Me.Range("A2").value = "Advanced User" Then
        ReDim EvTypeStr(0 To 3)
        EvTypeStr(0) = "ACTUAL"
        EvTypeStr(1) = "SCENARIO"
        EvTypeStr(2) = "HEARTBEAT"
        EvTypeStr(3) = "ALL"
        
    Else
        ReDim EvTypeStr(0 To 2)
        EvTypeStr(0) = "ACTUAL"
        EvTypeStr(1) = "SCENARIO"
        EvTypeStr(2) = "HEARTBEAT"
    End If
    
    With EvType.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Join(EvTypeStr, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Event Type"
        .ErrorTitle = ""
        .InputMessage = "Please select an event type from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
NotFormatDrop:
    If colNum <> 8 Then GoTo ExitHandler
    
    Set NotForm = Me.Range("H" & (rowNum))
    
    ' create an array for notification format
    Dim NotForms(0 To 2) As String
    NotForms(0) = "Rich Content"
    NotForms(1) = "Plain Text"
    NotForms(2) = "PAGER"
    
    With NotForm.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=Join(NotForms, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Notification Format"
        .ErrorTitle = ""
        .InputMessage = "Please select a notification format from the drop-down list"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
ExitHandler:
    Application.EnableEvents = True
    Application.Run "ProtectWorkbook"
End Sub









