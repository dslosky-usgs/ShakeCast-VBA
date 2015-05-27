VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DialogueForm 
   Caption         =   "Worksheet Dialogue"
   ClientHeight    =   8280.001
   ClientLeft      =   -40
   ClientTop       =   -2840
   ClientWidth     =   10360
   OleObjectBlob   =   "DialogueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DialogueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub ScrollBar1_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub DialogueBox_Change()

End Sub

Private Sub DialogueSave_Click()
Close #3

' Create a string that contains the time. This stops us from overwriting dialogue files
Dim timeStr As String
timeStr = Format(Now(), "yyyyMMdd_hh_mm_ss")

' Determine the computer's architecture so we know what delimeter to use to save the file
Dim getOS As String
getOS = Application.OperatingSystem

' Get the path that the workbook is being operated out of
Dim dir As String
dir = Application.ActiveWorkbook.Path

DiaExportForm.FileDest.Text = dir
DiaExportForm.FileName.Text = "Dialogue_" & timeStr & ".txt"
DiaExportForm.Show

' create the file name
If InStr(getOS, "Windows") = 0 Then
    Diapath = DiaExportForm.FileDest.Text & ":" & DiaExportForm.FileName.Text
Else
    Diapath = DiaExportForm.FileDest.Text & "\" & DiaExportForm.FileName.Text
End If

Open Diapath For Output As #3

Print #3, Me.DialogueBox.Text

If Diapath = ":" Or Diapath = "\" Then
    MsgBox "The dialogue was not saved"
Else
    MsgBox "Your dialogue was saved as " & vbNewLine & vbNewLine & _
    Diapath
End If
    
Close #3
    

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub DialogueFinished_Click()

DialogueForm.Hide


End Sub

Private Sub UserForm_Activate()

    Me.DialogueBox.SetFocus
    Me.DialogueBox.CurLine = 0
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
End Sub
