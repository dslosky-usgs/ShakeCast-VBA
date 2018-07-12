VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AttForm 
   Caption         =   "Facility Attributes"
   ClientHeight    =   5120
   ClientLeft      =   -1400
   ClientTop       =   -10660
   ClientWidth     =   8040
   OleObjectBlob   =   "AttForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AttForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Private Sub AddButton_Click()

    AttCheckBox.Show

End Sub

Private Sub AttFrame_Click()

End Sub

Private Sub CancelButton_Click()
    
    Unload Me
    

End Sub

Private Sub ManageButton_Click()

    ManageAtts.Show

End Sub

Private Sub OkayButton_Click()

Dim attStr As String
attStr = ""

For Each cont In AttFrame.Controls

    If attStr = "" Then
        attStr = cont.Caption
    ElseIf InStr(cont.Name, "Check") Then
        attStr = attStr & "%" & cont.Caption
    Else
        attStr = attStr & ":" & cont.Text
        
        If InStr(cont.Text, "%") Then
            MsgBox "Sorry: ""&"" is an invalid character"
            cont.Text = ""
            Exit Sub
        End If
        
    End If

Next cont

ActiveCell.value = attStr
    
Unload Me
End Sub

Private Sub CheckScroll()

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

Private Sub UserForm_Initialize()
If IsEmpty(ActiveCell) Then Exit Sub

Set myFrame = Me.AttFrame

Dim attStr As String
Dim attArr() As String
Dim eachAtt() As String
Dim i As Integer

attStr = ActiveCell.value
attArr = Split(attStr, "%")
i = 0

' fill both
For Each att In attArr

    eachAtt = Split(att, ":")
    
    att_name = eachAtt(0)
    eachAtt(0) = ""
    att_val = Join(eachAtt, ":")
    att_val = Right(att_val, Len(att_val) - 1)
    
    Set lab = myFrame.Controls.add("Forms.Label.1", "CheckBox_" & i)
    lab.Caption = att_name
    lab.Left = 5
    lab.Top = 5 + (i * 28)
    lab.Font.Size = 12
    lab.Height = 22
            
    Set txtbox = myFrame.Controls.add("Forms.TextBox.1", att_name)
    txtbox.Text = att_val
    txtbox.Left = 150
    txtbox.Top = 5 + (i * 28)
    txtbox.Font.Size = 12
    txtbox.Height = 22
    i = i + 1
    

Next att

Dim totalHeight As Integer
totalHeight = 5

For Each cont In Me.AttFrame.Controls

    If TypeOf cont Is MSForms.Label Then
        totalHeight = totalHeight + 28
    End If

Next cont

If totalHeight > Me.AttFrame.Height Then
    Me.AttFrame.ScrollHeight = totalHeight
End If

End Sub

