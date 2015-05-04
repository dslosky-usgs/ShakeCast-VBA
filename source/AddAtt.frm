VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddAtt 
   Caption         =   "Add an Attribute"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   OleObjectBlob   =   "AddAtt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddAtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub CreateButton_Click()

If Me.Caption = "Add an Attribute" Then

    Set mySheet = Worksheets("ShakeCast Ref Lookup Values")
    ' get attribute string
    Dim attStr As String
    attStr = mySheet.Range("P2").Value
    ' add value to it
    Dim newAttStr As String
    
    If attStr = "" Then
        newAttStr = AttName.Text
    Else
        newAttStr = attStr & "%" & AttName.Text
    End If
    
    If attStr <> "" Then
        If InArray(Split(attStr, "%"), AttName.Text) Then
            MsgBox "This attribute already exists!"
            GoTo TheEnd
        End If
    End If
    
    mySheet.Range("P2").Value = newAttStr
    
    ' refresh the ManageAtts form
    
    ' delete in text box
    For Each cont In ManageAtts.AttFrame.Controls
        If TypeOf cont Is MSForms.CheckBox Then
        'Cont.Delete
            ManageAtts.AttFrame.Controls.Remove cont.Name
        End If
    Next cont
    ' load attributes in text box
    
    Set attCell = mySheet.Range("P2")
    
    Dim attList() As String
    
    attList = Split(attCell.Value, "%")
    
    'For Each attCell In attCells
    '
    '    If Not IsEmpty(attCell) Then
    '        ReDim attList(0 To attCount)
    '
    '        attList(attCount) = attCell.Value
    '
    '        attCount = attCount + 1
    '    End If
    '
    'Next attCell
    
    ' turn list into check boxes
    Dim curColumn   As Long
    Dim lastRow     As Long
    Dim i           As Long
    Dim chkBox      As MSForms.CheckBox
    Dim subPos      As Integer
    
    subPos = 0
    For i = 0 To UBound(attList)
        If attList(i) = "" Then
            subPos = subPos + 1
            GoTo Nexti
        End If
    
        Set chkBox = ManageAtts.AttFrame.Add("Forms.CheckBox.1", "CheckBox_" & i)
        chkBox.Caption = attList(i)
        chkBox.Left = 5
        chkBox.Top = 5 + (i * 28) - (subPos * 28)
        chkBox.Font.Size = 12
        chkBox.Height = 22
        
        ' select the right checkboxes
    '    If InStr(ActiveCell.Value, GroupNames(i)) Then
    '        chkBox.Value = True
    '    End If
    
Nexti:
    Next i
    
    
    
    Dim totalHeight As Integer
    totalHeight = 5
    
    For Each cont In ManageAtts.AttFrame.Controls
    
        If TypeOf cont Is MSForms.CheckBox Then
            totalHeight = totalHeight + 28
        End If
    
    Next cont
    
    If totalHeight > ManageAtts.AttFrame.Height Then
        ManageAtts.AttFrame.ScrollHeight = totalHeight
    End If

ElseIf Me.Caption = "Create a Facility Type" Then

    Set mySheet = Worksheets("ShakeCast Ref Lookup Values")

    lastRow = mySheet.Cells(Rows.count, "C").End(xlUp).row + 1
    
    mySheet.Range("C" & lastRow).Value = AttName.Text
End If

Unload Me

TheEnd:
End Sub

Private Sub UserForm_Click()

End Sub
