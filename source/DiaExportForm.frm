VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DiaExportForm 
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   OleObjectBlob   =   "DiaExportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DiaExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ContDiaExport_Click()

On Error Resume Next

Dim dir As String
dir = ExportXML.FileDest.Text

MkDir dir

Err.Clear

DiaExportForm.Hide

End Sub

Private Sub SaveDialogue_Click()

Application.Run ("ExportDialogue")

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub
