VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Progress"
   ClientHeight    =   1720
   ClientLeft      =   40
   ClientTop       =   -2380
   ClientWidth     =   5680
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Frame1_Click()

End Sub

Private Sub ProcessName_Click()

End Sub

Private Sub ProgressLabel_Click()

End Sub

Private Sub UserForm_Activate()
    
    ProgressForm.ProgressLabel.Width = 0
        
    Set processCell = Worksheets("ShakeCast Ref Lookup Values").Range("Q2")

    Dim process As String
    process = processCell.value

    If process = "FacilityXML" Then
    
        ProgressForm.ProcessName.Caption = "Make Facility XML Table"
        Application.Run "FacilityXMLButton"
        
    ElseIf process = "GroupXML" Then
    
        ProgressForm.ProcessName.Caption = "Make Group XML Table"
        Application.Run "GroupXMLButton"
        
    ElseIf process = "UserXML" Then
    
        ProgressForm.ProcessName.Caption = "Make User XML Table"
        Application.Run "UserXMLButton"
        
    ElseIf process = "MasterXML" Then
    
        ProgressForm.ProcessName.Caption = "Make User XML Table"
        Application.Run "masterXMLexport"
    ElseIf process = "FacUpdate" Then
        ProgressForm.ProcessName.Caption = "Updating Worksheet"
        Application.Run "UpdateFacButton"
        
    ElseIf process = "UserUpdate" Then
        ProgressForm.ProcessName.Caption = "Updating Worksheet"
        Application.Run "UpdateGroupsButton"

    ElseIf process = "ImportCSV" Then
        ProgressForm.ProcessName.Caption = "Importing CSV"
        Application.Run "importCSV"
        
    End If

    processCell.value = Empty

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    

End Sub
