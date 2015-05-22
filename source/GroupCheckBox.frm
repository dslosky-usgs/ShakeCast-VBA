VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GroupCheckBox 
   Caption         =   "Group Selection"
   ClientHeight    =   5880
   ClientLeft      =   40
   ClientTop       =   -2840
   ClientWidth     =   4720
   OleObjectBlob   =   "GroupCheckBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GroupCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub AllButton_Click()
' select all the group names
For i = 0 To Me.GroupCount - 1

    Me.Controls("CheckBox_" & i).value = True
    
Next i

End Sub

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub ClearButton_Click()
' un-select all the group names

For i = 0 To Me.GroupCount - 1

    Me.Controls("CheckBox_" & i).value = False
    
Next i

End Sub

Private Sub OkayButton_Click()
' Turn selected group names into a string

Dim groupStr As String
groupStr = ""

For i = 0 To Me.GroupCount - 1
    If Me.Controls("CheckBox_" & i).value = True Then
        If groupStr = "" Then
            groupStr = Me.Controls("CheckBox_" & i).Caption
        Else
            groupStr = groupStr & ":" & Me.Controls("CheckBox_" & i).Caption
        End If
    End If
Next i

If groupStr = "" Then
    ActiveCell.value = Empty
Else
    ActiveCell.value = groupStr
End If

Unload Me


End Sub

Private Sub UserForm_Activate()




End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Application.Run "makeGroupChecklist"

End Sub
