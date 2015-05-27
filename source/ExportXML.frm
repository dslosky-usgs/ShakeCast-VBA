VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportXML 
   ClientHeight    =   3600
   ClientLeft      =   -40
   ClientTop       =   -2840
   ClientWidth     =   7300
   OleObjectBlob   =   "ExportXML.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ContXMLExport_Click()

On Error Resume Next

Dim dir As String
dir = ExportXML.FileDest.Text

MkDir dir

Err.Clear

Me.Hide

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
