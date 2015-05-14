VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageAtts 
   Caption         =   "Manage Attributes"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   OleObjectBlob   =   "ManageAtts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageAtts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CloseButton_Click()

    Unload Me

End Sub

Private Sub CreateButton_Click()

    AddAtt.Show


End Sub

Private Sub DeleteButton_Click()

Set mySheet = Worksheets("ShakeCast Ref Lookup Values")

' make a list of attributes to delete
Dim deleteAtt() As String
Dim attCount As Integer

attCount = 0
For Each Control In Me.AttFrame.Controls
    If TypeOf Control Is MSForms.CheckBox Then
        If Control.value = True Then
            ReDim Preserve deleteAtt(0 To attCount)
            deleteAtt(attCount) = Control.Caption
            attCount = attCount + 1
        End If
    End If
Next Control



' edit active cell str
Dim attStr As String
Dim newAttStr As String
Dim attArr() As String
Dim eachAtt() As String
attStr = mySheet.Range("P2").value
newAttStr = ""

attArr = Split(attStr, "%")


For Each entry In attArr

    If Not InArray(deleteAtt, entry) Then
        If newAttStr = "" Then
            newAttStr = entry
        Else
            newAttStr = newAttStr & "%" & entry
        End If
    End If
    
Next entry


mySheet.Range("P2").value = newAttStr


' delete in text box
For Each cont In Me.AttFrame.Controls
    If TypeOf cont Is MSForms.CheckBox Then
    'Cont.Delete
        Me.AttFrame.Controls.Remove cont.Name
    End If
Next cont
' load attributes in text box

Set attCell = mySheet.Range("P2")

Dim attList() As String

attList = Split(attCell.value, "%")

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

For i = 0 To UBound(attList)
    Set chkBox = Me.AttFrame.Add("Forms.CheckBox.1", "CheckBox_" & i)
    chkBox.Caption = attList(i)
    chkBox.Left = 5
    chkBox.Top = 5 + (i * 20)
    chkBox.Font.Size = 12
    
    
    ' select the right checkboxes
'    If InStr(ActiveCell.Value, GroupNames(i)) Then
'        chkBox.Value = True
'    End If
Next i


Dim totalHeight As Integer
totalHeight = 5

For Each cont In Me.AttFrame.Controls

    If TypeOf cont Is MSForms.Label Then
        totalHeight = totalHeight + 20
    End If

Next cont

If totalHeight > Me.AttFrame.Height Then
    Me.AttFrame.Height = totalHeight + 20
End If

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Set mySheet = Worksheets("ShakeCast Ref Lookup Values")

' get rid of any checkboxes that currently exist
For Each Control In Me.AttFrame.Controls
    If TypeOf Control Is MSForms.CheckBox Then
    Control.Delete
    End If
Next Control

Set attCell = mySheet.Range("P2")

Dim attList() As String
Dim attCount As Integer

attList = Split(attCell.value, "%")

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

For i = 0 To UBound(attList)
    Set chkBox = Me.AttFrame.Add("Forms.CheckBox.1", "CheckBox_" & i)
    chkBox.Caption = attList(i)
    chkBox.Left = 5
    chkBox.Top = 5 + (i * 28)
    chkBox.Font.Size = 12
    chkBox.Height = 22
    
    ' select the right checkboxes
'    If InStr(ActiveCell.Value, GroupNames(i)) Then
'        chkBox.Value = True
'    End If
Next i


' Keep the number of check boxes we've created to reference later in a hidden label

Dim totalHeight As Integer
totalHeight = 5

For Each cont In Me.AttFrame.Controls

    If TypeOf cont Is MSForms.CheckBox Then
        totalHeight = totalHeight + 28
    End If

Next cont

If totalHeight > Me.AttFrame.Height Then
    Me.AttFrame.ScrollHeight = totalHeight
End If

End Sub
