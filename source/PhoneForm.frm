VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhoneForm 
   Caption         =   "Enter a Phone Number"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   OleObjectBlob   =   "PhoneForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhoneForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ACText_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ACText_Change()

    ' make sure that the user has entered a number
    Dim newNum As String
    newNum = ""
    For i = 1 To Len(Me.ACText.Text)
        If InStr("0123456789", Mid(Me.ACText.Text, i, 1)) = 0 Then
            Me.ACText.Text = newNum
            MsgBox "That is not a number!"
            
            Exit Sub
        Else
            newNum = newNum & Mid(Me.ACText.Text, i, 1)
        End If
    Next i

    If Len(Me.ACText.Text) = 3 Then
        Me.bDash.SetFocus
    End If
    

End Sub

Private Sub aDash_Change()

    ' make sure that the user has entered a number
    Dim newNum As String
    newNum = ""
    For i = 1 To Len(Me.aDash.Text)
        If InStr("0123456789", Mid(Me.aDash.Text, i, 1)) = 0 Then
            Me.aDash.Text = newNum
            MsgBox "That is not a number!"
            
            Exit Sub
        Else
            newNum = newNum & Mid(Me.aDash.Text, i, 1)
        End If
    Next i

    If Len(Me.aDash.Text) = 4 Then
        Me.OkayButton.SetFocus
    End If

End Sub

Private Sub bDash_Change()

    ' make sure that the user has entered a number
    Dim newNum As String
    newNum = ""
    For i = 1 To Len(Me.bDash.Text)
        If InStr("0123456789", Mid(Me.bDash.Text, i, 1)) = 0 Then
            Me.bDash.Text = newNum
            MsgBox "That is not a number!"
            
            Exit Sub
        Else
            newNum = newNum & Mid(Me.bDash.Text, i, 1)
        End If
    Next i

    If Len(Me.bDash.Text) = 3 Then
        Me.aDash.SetFocus
    End If

End Sub

Private Sub OkayButton_Click()

    If Me.ACText.Text = "" And Me.bDash.Text = "" And Me.aDash.Text = "" Then
        ActiveCell.Value = Empty
        Unload Me
    ElseIf Me.ACText.Text = "" Or Me.bDash.Text = "" Or Me.aDash.Text = "" Then
        MsgBox "Invalid Phone Number"
        
    Else
    
        ActiveCell.Value = "(" & Me.ACText.Text & ")" & " " & Me.bDash.Text & "-" & Me.aDash.Text
        Unload Me
    End If
    
   

End Sub
