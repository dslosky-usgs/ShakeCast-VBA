VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} csvImport 
   ClientHeight    =   2260
   ClientLeft      =   40
   ClientTop       =   -2840
   ClientWidth     =   5500
   OleObjectBlob   =   "csvImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "csvImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Import_Click()

    Dim csvFile As String
    csvFile = Me.csv.Text
    
    If Right(csvFile, 4) <> ".csv" And Right(csvFile, 4) <> ".CSV" And Right(csvFile, 4) <> ".txt" And Right(csvFile, 4) <> ".TXT" Then
    
        MsgBox "This doesn't look like a csv file... check the file extension!"
        GoTo ExitHandler
        
    Else
        
        Unload Me
        Application.Run "importCSV", csvFile
        
    End If

ExitHandler:

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    Me.csv.Text = Application.ActiveWorkbook.Path
    

End Sub
